VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Shipping"
   ClientHeight    =   12810
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13545
   Icon            =   "shipping1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "shipping1.frx":030A
   ScaleHeight     =   12810
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   0
      TabIndex        =   53
      Top             =   2040
      Width           =   11175
   End
   Begin VB.TextBox sycranedb 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   52
      Text            =   "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
      Top             =   5640
      Width           =   8895
   End
   Begin VB.TextBox fmtfile 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   47
      Text            =   "S:\wd\bin\labfmt.txt"
      Top             =   6360
      Width           =   8895
   End
   Begin VB.TextBox srserv 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   42
      Top             =   6120
      Width           =   8895
   End
   Begin VB.TextBox pallogs 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   41
      Top             =   5880
      Width           =   8895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   40
      Top             =   6960
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid hgrid 
      Height          =   2175
      Left            =   0
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   37
      Top             =   7800
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4683
      _Version        =   327680
      BackColorFixed  =   65535
   End
   Begin VB.TextBox sybbsr 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   36
      Top             =   5400
      Width           =   8895
   End
   Begin VB.TextBox syship 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   35
      Top             =   5160
      Width           =   8895
   End
   Begin VB.TextBox babbsr 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   34
      Top             =   4920
      Width           =   8895
   End
   Begin VB.TextBox baship 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   33
      Top             =   4680
      Width           =   8895
   End
   Begin VB.TextBox plantno 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   28
      Top             =   3480
      Width           =   8895
   End
   Begin VB.TextBox ratrls 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   27
      Top             =   3240
      Width           =   8895
   End
   Begin VB.ListBox List2 
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   9360
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   10200
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox drvdir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Top             =   4440
      Width           =   8895
   End
   Begin VB.TextBox trltrk 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   4200
      Width           =   8895
   End
   Begin VB.TextBox webdir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   3960
      Width           =   8895
   End
   Begin MSFlexGridLib.MSFlexGrid FrmGrid 
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      Top             =   8400
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7858
      _Version        =   327680
      Cols            =   5
      BackColorFixed  =   12648384
   End
   Begin VB.TextBox ftpdir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Width           =   8895
   End
   Begin VB.TextBox bbsr 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   3000
      Width           =   8895
   End
   Begin VB.TextBox schdb 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   8895
   End
   Begin VB.TextBox shipdb 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   2520
      Width           =   8895
   End
   Begin VB.TextBox repdir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   8895
   End
   Begin VB.TextBox tempdir 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   8895
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SY Crane DB:"
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
      Left            =   0
      TabIndex        =   51
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      TabIndex        =   50
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   0
      TabIndex        =   49
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   0
      TabIndex        =   48
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label Pics"
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
      Left            =   0
      TabIndex        =   46
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label userid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   45
      Top             =   6600
      Width           =   8895
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1350
      Left            =   3600
      Picture         =   "shipping1.frx":1F0E
      Top             =   480
      Width           =   7560
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR Server:"
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
      Left            =   0
      TabIndex        =   44
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pallet Logs:"
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
      Left            =   0
      TabIndex        =   43
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label hcolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "hcolor"
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SY Racks DB:"
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
      Left            =   0
      TabIndex        =   32
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SY Ship DB:"
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
      Left            =   0
      TabIndex        =   31
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BA Racks DB:"
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
      Left            =   0
      TabIndex        =   30
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BA Ship DB:"
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
      Left            =   0
      TabIndex        =   29
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plant #:"
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
      Left            =   0
      TabIndex        =   26
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RA Trailer File:"
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
      Left            =   0
      TabIndex        =   25
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2024.12.16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   1800
      Width           =   7575
   End
   Begin VB.Label cgrp 
      Caption         =   "cgrp"
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label cdate 
      Caption         =   "cdate"
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Directions:"
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
      Left            =   0
      TabIndex        =   17
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trailer Tracking:"
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
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Web Directory:"
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
      Left            =   0
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FTP Scripts:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ship DB:"
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
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reports Directory:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu xitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu ordmenu 
      Caption         =   "&Orders"
      Begin VB.Menu trnsched 
         Caption         =   "Transport Schedule"
      End
      Begin VB.Menu trnschnotes 
         Caption         =   "Transport Schedule - Notes"
      End
      Begin VB.Menu edbrorders 
         Caption         =   "Branch Orders"
      End
      Begin VB.Menu pparts 
         Caption         =   "Print Partial Lists"
      End
      Begin VB.Menu partpalp 
         Caption         =   "Partial Pallet Labels"
      End
      Begin VB.Menu edwhstotals 
         Caption         =   "Inventory Adjustments"
      End
      Begin VB.Menu singmat 
         Caption         =   "View Single Pallet Matches"
      End
      Begin VB.Menu drplist 
         Caption         =   "View Drop Items In Orders"
      End
      Begin VB.Menu pjobtrl 
         Caption         =   "Post Jobbing Orders to Trailers"
      End
      Begin VB.Menu procjob 
         Caption         =   "Process Jobbing Pallets"
      End
      Begin VB.Menu bobtotrl 
         Caption         =   "Bobtails, Parlor, FedEx, QC Removal"
      End
      Begin VB.Menu coneords 
         Caption         =   "Cone Orders"
      End
      Begin VB.Menu clrbrords 
         Caption         =   "Clear Branch Orders"
      End
   End
   Begin VB.Menu grpmenu 
      Caption         =   "Groups"
      Begin VB.Menu brgrpmen 
         Caption         =   "Post Branches to Groups"
      End
      Begin VB.Menu edgroups 
         Caption         =   "Edit Groups"
      End
      Begin VB.Menu grptotrl 
         Caption         =   "Post Groups To Trailers"
      End
   End
   Begin VB.Menu trlmenu 
      Caption         =   "Trailers"
      Begin VB.Menu trailbills 
         Caption         =   "Trailer Bills"
      End
      Begin VB.Menu edtrls 
         Caption         =   "Edit Trailer Orders"
      End
      Begin VB.Menu edbillc 
         Caption         =   "Bills of Lading"
      End
      Begin VB.Menu renamtrl 
         Caption         =   "Rename Trailer Order"
      End
      Begin VB.Menu clrtrls 
         Caption         =   "Clear Trailers"
      End
      Begin VB.Menu addtrl 
         Caption         =   "Add Trailer"
         Visible         =   0   'False
      End
      Begin VB.Menu tcycle 
         Caption         =   "Trailer Cycle"
         Visible         =   0   'False
      End
      Begin VB.Menu trlra 
         Caption         =   "Post To Trailer History"
      End
      Begin VB.Menu edtrlsht 
         Caption         =   "Edit Trailer Sheet"
      End
   End
   Begin VB.Menu netmenu 
      Caption         =   "Networks"
      Begin VB.Menu tcpmenu 
         Caption         =   "TCP/IP"
         Begin VB.Menu impsr 
            Caption         =   "Import Crane Inventory"
         End
         Begin VB.Menu impreg 
            Caption         =   "Import Regular Inventory"
         End
         Begin VB.Menu impba 
            Caption         =   "Import Broken Arrow Inventory"
         End
         Begin VB.Menu impsy 
            Caption         =   "Import Sylacauga Inventory"
         End
         Begin VB.Menu impbrorder 
            Caption         =   "Import Branch Orders"
         End
         Begin VB.Menu impsched 
            Caption         =   "Import Transport Schedule"
         End
         Begin VB.Menu sndship 
            Caption         =   "Send Shipping List"
         End
      End
      Begin VB.Menu wanmenu 
         Caption         =   "Wide Area Network"
         Begin VB.Menu wdbrowstat 
            Caption         =   "W/D Browser Status"
         End
         Begin VB.Menu homeupdt 
            Caption         =   "Homepage Updates"
         End
         Begin VB.Menu hpuserp 
            Caption         =   "View Home Page User Log"
         End
         Begin VB.Menu clrhplog 
            Caption         =   "Clear Home Page User Log"
         End
         Begin VB.Menu pro11iadj 
            Caption         =   "Process Branch Adjustments"
         End
         Begin VB.Menu gemmoh 
            Caption         =   "Prepare Oracle Inventory Reports"
         End
      End
   End
   Begin VB.Menu Repmenu 
      Caption         =   "Reports"
      Begin VB.Menu prtbbol 
         Caption         =   "Blank Bill of Lading"
      End
      Begin VB.Menu outstk 
         Caption         =   "Stock Sheets"
      End
      Begin VB.Menu brordprt 
         Caption         =   "Branch Orders - All"
      End
      Begin VB.Menu brord1 
         Caption         =   "Branch Order "
      End
      Begin VB.Menu bsolow 
         Caption         =   "Branch SKU Orders - Lowstock"
      End
      Begin VB.Menu prodtots 
         Caption         =   "Daily Production Totals"
      End
      Begin VB.Menu sporders 
         Caption         =   "Snack Plant Orders"
      End
      Begin VB.Menu plantotals 
         Caption         =   "Plant Pallet Totals"
      End
      Begin VB.Menu brnotes 
         Caption         =   "Branch Notes"
      End
      Begin VB.Menu skuordprt 
         Caption         =   "SKU - Orders"
      End
      Begin VB.Menu skuissprt 
         Caption         =   "SKU - Issues"
      End
      Begin VB.Menu whssum 
         Caption         =   "Warehouse Total Summary"
      End
      Begin VB.Menu eopwks 
         Caption         =   "E-O-P Worksheet"
      End
      Begin VB.Menu sealtrax 
         Caption         =   "Seal Tracking"
      End
      Begin VB.Menu opcount 
         Caption         =   "Order Pick Count Sheet"
      End
      Begin VB.Menu vuettot 
         Caption         =   "View Trailer Totals"
      End
   End
   Begin VB.Menu rctmenu 
      Caption         =   "Receipts"
      Begin VB.Menu gemmsched 
         Caption         =   "Import Oracle Production Schedule"
      End
      Begin VB.Menu pprodrct 
         Caption         =   "Process Tri Level Receipts"
      End
      Begin VB.Menu cantldate 
         Caption         =   "Cancel Date"
      End
      Begin VB.Menu instldate 
         Caption         =   "Insert Date"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "&Configure"
      Begin VB.Menu edbranch 
         Caption         =   "&Branches"
      End
      Begin VB.Menu edbrprod 
         Caption         =   "Branch Products"
      End
      Begin VB.Menu disxprod 
         Caption         =   "Discontinued Products"
      End
      Begin VB.Menu edjobbing 
         Caption         =   "Jobbing Accounts"
      End
      Begin VB.Menu edop 
         Caption         =   "Order Pick List"
      End
      Begin VB.Menu edplbranch 
         Caption         =   "Plant -> Branches"
      End
      Begin VB.Menu edplskus 
         Caption         =   "Plant -> Products"
      End
      Begin VB.Menu edplants 
         Caption         =   "Production Plants"
      End
      Begin VB.Menu edprsource 
         Caption         =   "Production Sources"
      End
      Begin VB.Menu edsku 
         Caption         =   "SKU Master"
      End
      Begin VB.Menu skucomp 
         Caption         =   "SKU Master - Comp"
      End
      Begin VB.Menu vallists 
         Caption         =   "Value Lists"
      End
      Begin VB.Menu edwhs 
         Caption         =   "Warehouses"
      End
      Begin VB.Menu edusers 
         Caption         =   "W-D Users"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function stringit(ss As String) As String
    Dim i As Integer, k As Integer, s As String
    k = Len(ss)
    If k = 0 Then
        s = "."
    Else
        s = ""
        For i = 1 To k
            s = s & mid(ss, i, 1) & " "
        Next i
    End If
    stringit = s
End Function

Function ellone(ss As String) As String
    Dim i As Integer, k As Integer, s As String
    k = Len(ss)
    If k = 0 Then
        s = "."
    Else
        s = ""
        For i = 1 To k
            If i = 1 Then
                s = s & mid(ss, i, 1)
            Else
                If mid(ss, i, 1) = "1" Then
                    s = s & "l"
                Else
                    s = s & mid(ss, i, 1)
                End If
            End If
        Next i
    End If
    ellone = s
End Function

Sub menu_build(uname As String)
    Dim ds As adodb.Recordset, s As String
    Me.addtrl.Visible = False
    Me.bobtotrl.Visible = False
    Me.brgrpmen.Visible = False
    Me.brnotes.Visible = False
    Me.brord1.Visible = False
    Me.brordprt.Visible = False
    Me.bsolow.Visible = False
    Me.cantldate.Visible = False
    Me.clrbrords.Enabled = False: Me.clrbrords.Caption = "..."
    Me.clrhplog.Visible = False
    Me.clrtrls.Visible = False
    Me.coneords.Visible = False
    Me.disxprod.Visible = False
    Me.drplist.Visible = False
    Me.edbillc.Visible = False
    Me.edbranch.Visible = False
    Me.edbrorders.Visible = False
    Me.edbrprod.Visible = False
    Me.edgroups.Visible = False
    Me.edjobbing.Visible = False
    Me.edmenu.Visible = False
    Me.edop.Visible = False
    Me.edplants.Visible = False
    Me.edplbranch.Visible = False
    Me.edplskus.Visible = False
    Me.edprsource.Visible = False
    Me.edsku.Visible = False
    Me.edtrls.Visible = False
    Me.edtrlsht.Enabled = False: Me.edtrlsht.Caption = "..."
    Me.edusers.Enabled = False: Me.edusers.Caption = "..."
    Me.edwhs.Visible = False
    Me.edwhstotals.Visible = False
    Me.eopwks.Visible = False
    Me.gemmoh.Enabled = False: Me.gemmoh.Caption = "..."
    Me.grpmenu.Visible = False
    Me.grptotrl.Enabled = False: Me.grptotrl.Caption = "..."
    Me.homeupdt.Visible = False
    Me.hpuserp.Visible = False
    Me.impba.Visible = False
    Me.impbrorder.Visible = False
    Me.impreg.Visible = False
    Me.impsched.Visible = False
    Me.impsr.Visible = False
    Me.impsy.Visible = False
    Me.instldate.Enabled = False: Me.instldate.Caption = "..."
    Me.netmenu.Visible = False
    Me.opcount.Visible = False
    Me.ordmenu.Visible = False
    Me.outstk.Visible = False
    Me.partpalp.Visible = False
    Me.pjobtrl.Visible = False
    Me.plantotals.Visible = False
    Me.pparts.Visible = False
    Me.pprodrct.Visible = False
    Me.pro11iadj.Visible = False
    Me.procjob.Visible = False
    Me.prodtots.Visible = False
    Me.prtbbol.Visible = False
    Me.rctmenu.Visible = False
    'Me.renamtrl.Visible = False
    Me.renamtrl.Enabled = False
    Me.Repmenu.Visible = False
    Me.singmat.Visible = False
    Me.skucomp.Visible = False
    Me.skuissprt.Visible = False
    Me.skuordprt.Visible = False
    Me.sndship.Enabled = False: Me.sndship.Caption = "..."
    Me.sporders.Visible = False
    Me.tcpmenu.Visible = False
    Me.tcycle.Visible = False
    Me.trailbills.Visible = False
    Me.trlmenu.Visible = False
    Me.trlra.Visible = False
    Me.trnsched.Visible = False
    Me.trnschnotes.Visible = False
    Me.vallists.Visible = False
    Me.vuettot.Enabled = False: Me.vuettot.Caption = "..."
    Me.wanmenu.Enabled = False: Me.wanmenu.Caption = "..."
    Me.wdbrowstat.Visible = False
    Me.whssum.Visible = False
    
    s = "select menuname from usermenus where userid = '" & uname & "'"
    If Form1.plantno = "50" Then s = s & " and orgid = '500'"
    If Form1.plantno = "51" Then s = s & " and orgid = '501'"
    If Form1.plantno = "52" Then s = s & " and orgid = '502'"
    Set ds = Sdb.Execute(s)         'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!menuname = "addtrl" Then Me.addtrl.Visible = True
            If ds!menuname = "bobtotrl" Then Me.bobtotrl.Visible = True
            If ds!menuname = "brgrpmen" Then Me.brgrpmen.Visible = True
            If ds!menuname = "brnotes" Then Me.brnotes.Visible = True
            If ds!menuname = "brord1" Then Me.brord1.Visible = True
            If ds!menuname = "brordprt" Then Me.brordprt.Visible = True
            If ds!menuname = "bsolow" Then Me.bsolow.Visible = True
            If ds!menuname = "cantldate" Then Me.cantldate.Visible = True
            If ds!menuname = "clrbrords" Then
                Me.clrbrords.Enabled = True: Me.clrbrords.Caption = "Clear Branch Orders"
            End If
            If ds!menuname = "clrhplog" Then Me.clrhplog.Visible = True
            If ds!menuname = "clrtrls" Then Me.clrtrls.Visible = True
            If ds!menuname = "coneords" Then Me.coneords.Visible = True
            If ds!menuname = "disxprod" Then Me.disxprod.Visible = True
            If ds!menuname = "drplist" Then Me.drplist.Visible = True
            If ds!menuname = "edbillc" Then Me.edbillc.Visible = True
            If ds!menuname = "edbranch" Then Me.edbranch.Visible = True
            If ds!menuname = "edbrorders" Then Me.edbrorders.Visible = True
            If ds!menuname = "edbrprod" Then Me.edbrprod.Visible = True
            If ds!menuname = "edgroups" Then Me.edgroups.Visible = True
            If ds!menuname = "edjobbing" Then Me.edjobbing.Visible = True
            If ds!menuname = "edmenu" Then Me.edmenu.Visible = True
            If ds!menuname = "edop" Then Me.edop.Visible = True
            If ds!menuname = "edplants" Then Me.edplants.Visible = True
            If ds!menuname = "edplbranch" Then Me.edplbranch.Visible = True
            If ds!menuname = "edplskus" Then Me.edplskus.Visible = True
            If ds!menuname = "edprsource" Then Me.edprsource.Visible = True
            If ds!menuname = "edsku" Then Me.edsku.Visible = True
            If ds!menuname = "edtrls" Then Me.edtrls.Visible = True
            If ds!menuname = "edtrlsht" Then
                Me.edtrlsht.Enabled = True: Me.edtrlsht.Caption = "Edit Trailer Sheet"
            End If
            If ds!menuname = "edusers" Then
                Me.edusers.Enabled = True: Me.edusers.Caption = "W-D Users"
            End If
            If ds!menuname = "edwhs" Then Me.edwhs.Visible = True
            If ds!menuname = "edwhstotals" Then Me.edwhstotals.Visible = True
            If ds!menuname = "eopwks" Then Me.eopwks.Visible = True
            If ds!menuname = "gemmoh" Then
                Me.gemmoh.Enabled = True: Me.gemmoh.Caption = "Prepare Oracle Inventory Reports"
            End If
            If ds!menuname = "grpmenu" Then Me.grpmenu.Visible = True
            If ds!menuname = "grptotrl" Then
                Me.grptotrl.Enabled = True: Me.grptotrl.Caption = "Post Groups To Trailers"
            End If
            If ds!menuname = "homeupdt" Then Me.homeupdt.Visible = True
            If ds!menuname = "hpuserp" Then Me.hpuserp.Visible = True
            If ds!menuname = "impba" Then Me.impba.Visible = True
            If ds!menuname = "impbrorder" Then Me.impbrorder.Visible = True
            If ds!menuname = "impreg" Then Me.impreg.Visible = True
            If ds!menuname = "impsched" Then Me.impsched.Visible = True
            If ds!menuname = "impsr" Then Me.impsr.Visible = True
            If ds!menuname = "impsy" Then Me.impsy.Visible = True
            If ds!menuname = "instldate" Then
                Me.instldate.Enabled = True: Me.instldate.Caption = "Insert Date"
            End If
            If ds!menuname = "netmenu" Then Me.netmenu.Visible = True
            If ds!menuname = "opcount" Then Me.opcount.Visible = True
            If ds!menuname = "ordmenu" Then Me.ordmenu.Visible = True
            If ds!menuname = "outstk" Then Me.outstk.Visible = True
            If ds!menuname = "partpalp" Then Me.partpalp.Visible = True
            If ds!menuname = "pjobtrl" Then Me.pjobtrl.Visible = True
            If ds!menuname = "plantotals" Then Me.plantotals.Visible = True
            If ds!menuname = "pparts" Then Me.pparts.Visible = True
            If ds!menuname = "pprodrct" Then Me.pprodrct.Visible = True
            If ds!menuname = "pro11iadj" Then Me.pro11iadj.Visible = True
            If ds!menuname = "procjob" Then Me.procjob.Visible = True
            If ds!menuname = "prodtots" Then Me.prodtots.Visible = True
            If ds!menuname = "prtbbol" Then Me.prtbbol.Visible = True
            If ds!menuname = "rctmenu" Then Me.rctmenu.Visible = True
            If ds!menuname = "renamtrl" Then Me.renamtrl.Enabled = True
            If ds!menuname = "Repmenu" Then Me.Repmenu.Visible = True
            If ds!menuname = "singmat" Then Me.singmat.Visible = True
            If ds!menuname = "skucomp" Then Me.skucomp.Visible = True
            If ds!menuname = "skuissprt" Then Me.skuissprt.Visible = True
            If ds!menuname = "skuordprt" Then Me.skuordprt.Visible = True
            If ds!menuname = "sndship" Then
                Me.sndship.Enabled = True: Me.sndship.Caption = "Send Shipping Orders"
            End If
            If ds!menuname = "sporders" Then Me.sporders.Visible = True
            If ds!menuname = "tcpmenu" Then Me.tcpmenu.Visible = True
            If ds!menuname = "tcycle" Then Me.tcycle.Visible = True
            If ds!menuname = "trailbills" Then Me.trailbills.Visible = True
            If ds!menuname = "trlmenu" Then Me.trlmenu.Visible = True
            If ds!menuname = "trlra" Then Me.trlra.Visible = True
            If ds!menuname = "trnsched" Then Me.trnsched.Visible = True
            If ds!menuname = "trnschnotes" Then Me.trnschnotes.Visible = True
            If ds!menuname = "vallists" Then Me.vallists.Visible = True
            If ds!menuname = "vuettot" Then
                Me.vuettot.Enabled = True: Me.vuettot.Caption = "View Trailer Totals"
            End If
            If ds!menuname = "wanmenu" Then
                Me.wanmenu.Enabled = True: Me.wanmenu.Caption = "Wide Area Network"
            End If
            If ds!menuname = "wdbrowstat" Then Me.wdbrowstat.Visible = True
            If ds!menuname = "whssum" Then Me.whssum.Visible = True
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub schedlabels()               'jv061914
    Dim sqlx As String, pdate As String, punit As String
    Dim pflag As String, plot As String, psku As String
    Dim pcode As String, f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim ppal As Integer, pday As Integer, tl As Integer, pdaze As Integer
    Dim ds As adodb.Recordset, ds2 As adodb.Recordset, ds3 As adodb.Recordset, cflag As Integer
    Dim i, userid As String, pwd As String, dsn As String, query As String
    Dim scdate As String, zid As Long
    Dim cfile As String
    cfile = "S:\wd\data\plabels.500"
    If Len(Dir(cfile)) = 0 Then
        MsgBox "Label File: " & cfile & " is not found.", vbOKOnly + vbInformation, "sorry, try again.."
        Exit Sub
    End If
    'On Error Resume Next
    
    scdate = InputBox("Number of days from today", "Schedule Days", "1")
    If Len(scdate) = 0 Then
        Exit Sub
    End If
    Grid1.Clear: Grid1.Rows = 0: Grid1.Cols = 4
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        If Format(f1, "MM-dd-yyyy") = Format(DateAdd("d", scdate, Now), "MM-dd-yyyy") Then
            s = f2 & Chr(9) & f1 & Chr(9) & f4 & Chr(9) & f6
            Grid1.AddItem s
        End If
    Loop
    Close #1
    Grid1.FixedCols = 2
    Grid1.FormatString = "^sku|^DATE|^Units|^SP"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 1800
    Grid1.ColWidth(3) = 600
    pdate = Format(DateAdd("d", Val(scdate), Now), "M-d-yyyy")
    pday = 0: cflag = 1
    For k = 0 To Grid1.Rows - 1
        psku = Grid1.TextMatrix(k, 0)
        pdate = Grid1.TextMatrix(k, 1)
        punit = Grid1.TextMatrix(k, 2)
        pcode = Grid1.TextMatrix(k, 3)
        'If Val(psku) > 0 And Val(psku) < 1000 And IsDate(pdate) Then
        If Val(psku) > 0 And Val(psku) < 9999 And IsDate(pdate) Then            'jv062816
            If pday = 0 Then
                pday = 1
                If WeekDay(pdate) = 6 Then
                    If MsgBox("Production this Saturday?", vbYesNo + vbQuestion, "Friday Products") = vbYes Then
                        pday = 1
                    Else
                        pday = 3
                    End If
                End If
                If WeekDay(pdate) = 7 Then pday = 2
                plot = Right$(Format$(pdate, "m/d/yy"), 2)          'jv021615
                plot = plot & Format(DateDiff("d", DateValue("1/1/" & plot), DateValue(pdate)), "000")
                plot = Val(plot) + 1
            End If
            If cflag = 1 Then
                sqlx = "delete from prodrcv where proddate = '" & Format(pdate, "m-d-yyyy") & "'"
                Wdb.Execute sqlx                'jv060916
                cflag = 0
            End If
            s = "select * from skumast where sku = '" & psku & "'"
            Set ds = Sdb.Execute(s)         'jv060916
            If ds.BOF = True Then
                MsgBox "Invalid SKU: " & psku & "detected.", vbOKOnly, "Incoming garbage from Oracle.."
            Else
                ds.MoveFirst
                s = "select * from prodsources where source = " & ds!psource
                Set ds2 = Sdb.Execute(s)    'jv060916
                If ds2.BOF = True Then
                    tl = 0: pdaze = 1
                Else
                    ds2.MoveFirst
                    If ds2!tl_flag = "Y" Then       'jv011615
                        tl = 1                      'jv011615
                    Else                            'jv011615
                        tl = 0                      'jv011615
                    End If                          'jv011615
                    'tl = ds2!tl_flag:
                    pdaze = ds2!days
                End If
                ds2.Close
                ppal = punit
                punit = Format(Val(punit) * ds!pallet, "0")
                'ppal = Int((Val(punit) / Val(ds!pallet)) + 0.75)
                s = "select * from prodrcv where sku = '" & psku & "' and proddate = '" & Format(pdate, "m-d-yyyy") & "'"
                s = s & " and sp_flag = '" & pcode & "'"
                Set ds3 = Wdb.Execute(s)        'jv060916
                If ds3.BOF = True Then
                    zid = wd_seq("ProdRcv", Form1.bbsr)
                    s = "INSERT INTO ProdRcv (ID, SKU, ProdDate, Units, SP_Flag, Lot_Num,"
                    s = s & " RecDate1, RecDate2, RecDate3, SR1, SR2, SR3, SR4, SR5) VALUES (" & zid & ","
                    s = s & "'" & psku & "',"
                    s = s & "'" & pdate & "',"
                    s = s & punit & ","
                    's = s & "'" & Val(pflag) & "',"
                    s = s & "'" & pcode & "',"
                    s = s & "'" & Format(Val(plot), "00000") & "',"
                    s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
                    If pdaze > 1 Then
                        s = s & "'" & Format(DateAdd("d", pday, pdate), "mm-dd-yyyy") & "',"
                        s = s & "'" & Format(DateAdd("d", pday, pdate), "mm-dd-yyyy") & "',"
                    Else
                        s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
                        s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
                    End If
                    If ds!whs_num = 1 Then
                        s = s & ppal & ",0,0,0,0)"
                    Else
                        If ds!whs_num = 2 Then
                            s = s & "0," & ppal & ",0,0,0)"
                        Else
                            If ds!whs_num = 3 Then
                                s = s & "0,0," & ppal & ",0,0)"
                            Else
                                If ds!whs_num = 4 And tl <> 0 Then
                                    s = s & "0,0,0," & ppal & ",0)"
                                Else
                                    If ds!whs_num = 5 Then
                                        s = s & "0,0,0,0," & ppal & ")"
                                    Else
                                        s = s & "0,0,0,0,0)"
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Wdb.Execute s                   'jv060916
                Else
                    s = "Update prodrcv set units = units + " & punit
                    If ds!whs_num = 1 Then s = s & "'sr1 = sr1 + " & ppal
                    If ds!whs_num = 2 Then s = s & "'sr2 = sr2 + " & ppal
                    If ds!whs_num = 3 Then s = s & "'sr3 = sr3 + " & ppal
                    If ds!whs_num = 4 And tl <> 0 Then s = s & "'sr4 = sr4 + " & ppal
                    If ds!whs_num = 5 Then s = s & "'sr5 = sr5 + " & ppal
                    s = s & " where id = " & ds!id
                    Wdb.Execute s                   'jv060916
                End If
                ds3.Close
            End If
            ds.Close
        End If
    Next k
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Call pprodrct_Click
End Sub

Private Sub check_sku_orders()
    Dim ss As adodb.Recordset, ts As adodb.Recordset
    Dim ds As adodb.Recordset
    Dim s As String, oh As Long, pname As String, bname As String
    Dim rt As String, rf As String, rh As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    Screen.MousePointer = 11
    s = "select sku, sum(netqty) from brorders where plant = 50 and netqty > 0 group by sku"
    Set ss = Sdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            oh = 0
            s = "select sum(qty) from lane where sku = '" & ss!sku & "'"
            s = s & " having sum(qty) > 0"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                oh = oh + ds(0)
            End If
            ds.Close
            s = "select count(*) from rackpos where sku = '" & ss!sku & "'"
            s = s & " and rackno not in (select id from racks where rack = 'OP'"
            s = s & " or hold = 1)"
            s = s & " having count(*) > 0"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                oh = oh + ds(0)
            End If
            ds.Close
            s = "select count(*) from paltasks where area = 'DOCK'"
            s = s & " and product >= '" & ss!sku & "' and product <= '" & ss!sku & "ZZZ'"
            s = s & " and userid < '0'"
            s = s & " and status <> 'COMP'"
            s = s & " having count(*) > 0"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                oh = oh - ds(0)
            End If
            ds.Close
            If oh < ss(1) Then
                s = "select fgunit, fgdesc from skumast where sku = '" & ss!sku & "'"
                Set ts = Sdb.Execute(s)
                If ts.BOF = False Then
                    ts.MoveFirst
                    pdesc = ts!fgunit & " " & ts!fgdesc
                Else
                    pdesc = "unknown SKU"
                End If
                ts.Close
                s = ss!sku & Chr(9) & pdesc & Chr(9) & ss(1) & Chr(9) & oh
                Grid1.AddItem s
                s = "select brorders.branch, branchname, netqty from brorders,branches"
                s = s & " where sku = '" & ss!sku & "' and plant = 50 and netqty > 0"
                s = s & " and branches.branch = brorders.branch"
                Set ts = Sdb.Execute(s)
                If ts.BOF = False Then
                    ts.MoveFirst
                    Do Until ts.EOF
                        s = Chr(9) & Chr(9) & Chr(9) & Chr(9) & ts(0) & "-" & ts(1) & Chr(9) & ts(2)
                        Grid1.AddItem s
                        ts.MoveNext
                    Loop
                End If
                ts.Close
            End If
            ss.MoveNext
        Loop
    End If
    ss.Close
    Grid1.FormatString = "^SKU|<Product|^Orders|^Available|<Branch|^Qty"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 1000
    Screen.MousePointer = 0
    rt = "Branch SKU Orders"
    rh = Format(mdate, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        htdc(0) = "Yellow": gndc(0) = hcolor.BackColor
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "check_sku_orders", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check_sku_orders - Error Number: " & eno
        End
    End If
End Sub

Private Sub sylacauga_yard_countsheet()
    Dim ds As adodb.Recordset, s As String
    Dim i As Integer, mtag As String, stag As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim cfile As String, mdate As String
    'On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    cfile = "\\bbsy-02-dc\f\user\waredist\bin\transitin.502"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6 & Chr(9) & f7
            Grid1.AddItem s
        Loop
        Close #1
    End If
    stag = "XXX"
    s = "select shipdate,trlno,groupcode,trailers.sku,fgunit,fgdesc,sum(pallets),sum(wraps),sum(units)"
    s = s & " from trailers,skumast"
    s = s & " where shipdate = '" & Format(Now, "m-d-yyyy") & "'"
    s = s & " and plant in (50,51) and branch = 52"
    s = s & " and skumast.sku = trailers.sku"
    s = s & " group by shipdate,trlno,groupcode,trailers.sku,fgunit,fgdesc"
    s = s & " order by shipdate,trlno,trailers.sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            mtag = Format(ds!shipdate, "ddd") & " " & ds!trlno
            If stag <> mtag Then
                For i = Grid1.Rows - 1 To 1 Step -1
                    If Grid1.TextMatrix(i, 0) = mtag Then
                        If Grid1.Rows <= 2 Then
                            Grid1.Rows = 1
                        Else
                            Grid1.RemoveItem i
                        End If
                    End If
                Next i
                stag = mtag
            End If
            s = mtag & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & StrConv(ds!fgunit, vbProperCase) & " "
            s = s & StrConv(ds!fgdesc, vbProperCase) & Chr(9)
            s = s & Format(ds(6), "#") & Chr(9)
            s = s & Format(ds(7), "#") & Chr(9)
            s = s & Chr(9)
            s = s & Format(ds(8), "0") & Chr(9)
            s = s & Format(ds!shipdate, "m-d-yyyy")
            If LCase(Left(ds!groupcode, 2)) <> "oc" Then Grid1.AddItem s        'jv092408
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Tag|^SKU|<Description|^Pallets|^Wraps|^Units|^Total|^LastDate"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Grid1.TextMatrix(i, 0),
        Write #1, Grid1.TextMatrix(i, 1),
        Write #1, Grid1.TextMatrix(i, 2),
        Write #1, Grid1.TextMatrix(i, 3),
        Write #1, Grid1.TextMatrix(i, 4),
        Write #1, Grid1.TextMatrix(i, 5),
        Write #1, Grid1.TextMatrix(i, 6),
        Write #1, Grid1.TextMatrix(i, 7)
    Next i
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sylacauga_yard_countsheet", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sylacauga_yard_countsheet - Error Number: " & eno
        End
    End If
End Sub

Private Sub imp_sylacauga_lowstock()
    Dim ds As adodb.Recordset, sqlx As String
    Dim ss As adodb.Recordset, pkey As Long
    Dim query As String, i As Integer, k As Integer
    Dim dsn As String, userid As String, pwd As String
    On Error GoTo vberror
    Open Form1.srserv & "\wd\bin\gemmodbc.ini" For Input As #1
    Line Input #1, dsn
    Line Input #1, userid
    Line Input #1, pwd
    Close #1
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    Screen.MousePointer = 11
    'R12
    query = "select o.inventory_item_id, m.segment1, m.description, sum(o.transaction_quantity)" & _
            " from mtl_onhand_quantities o, mtl_system_items_b m" & _
            " where o.subinventory_code in ('A10','052')" & _
            " and m.organization_id = o.organization_id" & _
            " and m.inventory_item_id = o.inventory_item_id" & _
            " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
            " group by o.inventory_item_id, m.segment1, m.description" & _
            " order by m.segment1"

    i = LoadGrid(Grid1, query, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    
    Grid1.Cols = 6
    sqlx = "select sku,pallet from skumast where pallet in (60,468)"
    Set ds = Sdb.Execute(sqlx)              'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 1) = ds!sku Then
                    If Val(Grid1.TextMatrix(i, 3)) > 0 Then
                        Grid1.TextMatrix(i, 4) = Format(Val(Grid1.TextMatrix(i, 3)) / ds!pallet, ".000")
                    End If
                    sqlx = "select * from whstotals where whs_num = 15 and sku = '" & ds!sku & "'"
                    Set ss = Sdb.Execute(sqlx)  'jv060916
                    If ss.BOF = False Then
                        ss.MoveFirst
                        Grid1.TextMatrix(i, 5) = ss!avail
                    End If
                    ss.Close
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    For i = Grid1.Rows - 1 To 1 Step -1
        If Val(Grid1.TextMatrix(i, 4)) < 0.5 And Grid1.Rows > 2 Then
            Grid1.RemoveItem i
        Else
            k = 2
            If Val(Grid1.TextMatrix(i, 4)) >= 1.5 Then k = 2
            If Val(Grid1.TextMatrix(i, 4)) >= 2.5 Then k = 3
            If Val(Grid1.TextMatrix(i, 4)) >= 3.5 Then k = 4
            If Val(Grid1.TextMatrix(i, 4)) >= 4.5 Then k = 5
            k = CInt(Val(Grid1.TextMatrix(i, 4)) + 0.01)
            If k = 1 Then k = 2
            Grid1.TextMatrix(i, 5) = k
            s = "select * from whstotals where whs_num = 15 and sku = '" & Grid1.TextMatrix(i, 1) & "'"
            Set ds = Sdb.Execute(s)                 'jv060916
            If ds.BOF = False Then
                s = "Update whstotals set count_qty = " & Val(Grid1.TextMatrix(i, 5))
                s = s & ", avail = " & Val(Grid1.TextMatrix(i, 5)) & " - grp_qty where id = " & ds!id
            Else
                pkey = wd_seq("whstotals", Form1.shipdb)
                s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail, old_qty)"
                s = s & " Values (" & pkey & ", 15, '" & Grid1.TextMatrix(i, 1) & "'"
                s = s & ", " & Val(Grid1.TextMatrix(i, 5))
                s = s & ", 0, " & Val(Grid1.TextMatrix(i, 5)) & ", 0"
            End If
            ds.Close
            Sdb.Execute sqlx                        'jv060916
        End If
    Next i
    Grid1.FormatString = "^ID|^SKU|<Product|^Units|^Pallets|^Available"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "imp_sylacauga_lowstock", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_sylacauga_lowstock - Error Number: " & eno
        End
    End If
End Sub

Private Sub import_bc_racks()
    Dim d1 As adodb.Recordset, d2 As adodb.Recordset
    Dim d3 As adodb.Recordset, d4 As adodb.Recordset, pgrp As Integer, preg As Integer
    Dim prega As Integer, p4way As Integer, psp As Integer, s As String, pante As Integer, pkey As Long
    Screen.MousePointer = 11
    On Error GoTo vberror
    preg = 0: prega = 0: p4way = 0: psp = 0
    s = "select * from warehouses where plant = 50 and whs in ('REG', 'REGA', '4WAY', 'SP', 'ANTE')"
    Set d1 = Sdb.Execute(s)         'jv060916
    If d1.BOF = False Then
        d1.MoveFirst
        Do Until d1.EOF
            If UCase(d1!whs) = "REG" Then preg = d1!whs_num
            If UCase(d1!whs) = "REGA" Then prega = d1!whs_num
            If UCase(d1!whs) = "4WAY" Then p4way = d1!whs_num
            If UCase(d1!whs) = "SP" Then psp = d1!whs_num
            If UCase(d1!whs) = "ANTE" Then pante = d1!whs_num
            d1.MoveNext
        Loop
    End If
    d1.Close
    s = ""
    If preg = 0 Then s = s & "Did not detect REG warehouse in listing." & vbCrLf
    If prega = 0 Then s = s & "Did not detect REGA warehouse in listing." & vbCrLf
    If p4way = 0 Then s = s & "Did not detect 4WAY warehouse in listing." & vbCrLf
    If psp = 0 Then s = s & "Did not detect SP warehouse in listing." & vbCrLf
    If pante = 0 Then s = s & "Did not detect ANTE warehouse in listing." & vbCrLf
    If Len(s) > 0 Then
        Screen.MousePointer = 0
        MsgBox s, vbOKOnly + vbInformation, "aborting, check warehouse master table..."
        Exit Sub
    End If
    s = "Delete from whstotals where whs_num in (" & preg & "," & prega & "," & p4way & "," & psp & "," & pante & ")"
    Sdb.Execute s               'jv060916
    
    'Regular ------------------------------------------------------------------
    s = "select p.sku,r.fo,count(*) from racks r, rackpos p where p.rackno = r.id"
    s = s & " and r.aisle <> 'M'"
    s = s & " and r.hold = 0 and p.sku > '0000' and p.sku < '9999'"                 'jv082415
    s = s & " and p.bbc = 'Y'"                          'jv111611
    s = s & " group by p.sku,r.fo having count(*) > 0"
    Set d4 = Wdb.Execute(s)                 'jv060916
    If d4.BOF = False Then
        d4.MoveFirst
        Do Until d4.EOF
            s = "select * from skumast where sku = '" & Trim(Left(d4(0), 4)) & "'"          'jv082415
            Set d2 = Sdb.Execute(s)             'jv060916
            If d2.BOF = True Then
                MsgBox "What is this sku? " & d4(0), vbOKOnly, "Garbage found in BRW racks."
            Else
                s = "select product,count(*) from paltasks where area = 'FORKLIFT'"
                s = s & " and status = 'PEND' and target = 'STAGING'"
                s = s & " and product >= '" & d4(0) & "'"
                s = s & " and product < '" & d4(0) & "ZZZ'"
                s = s & " and source not in ('ANTE ROOM', 'SNACK PLANT')"
                s = s & " group by product having count(*) > 0"
                Set d3 = Wdb.Execute(s)         'jv060916
                If d3.BOF = False Then
                    d3.MoveFirst
                    pgrp = d3(1)
                Else
                    pgrp = 0
                End If
                d3.Close
                If d2!whs_num = preg Then
                    s = "select * from whstotals where sku = '" & Trim(Left(d4!sku, 4)) & "' and whs_num = " & preg     'jv082415
                    Set d1 = Sdb.Execute(s)         'jv060916
                    If d1.BOF = True Then
                        pkey = wd_seq("whstotals", Form1.shipdb)
                        s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                        s = s & ", " & preg & ", '" & Trim(Left(d4!sku, 4)) & "'"           'jv051016
                        s = s & ", " & d4(2) & ", " & pgrp & ", " & d4(2) - pgrp & ")"
                    Else
                        s = "Update whstotals set count_qty = count_qty + " & d4(2)
                        s = s & ", avail = avail + " & d4(2) & " Where id = " & d1!id
                    End If
                    Sdb.Execute s               'jv060916
                    d1.Close
                Else
                    pkey = wd_seq("whstotals", Form1.shipdb)
                    If d4!fo <> 0 Then
                        s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                        s = s & ", " & preg & ", '" & Trim(Left(d4!sku, 4)) & "'"           'jv082415
                        s = s & ", " & d4(2) & ", " & pgrp & ", " & d4(2) - pgrp & ")"
                    Else
                        s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                        s = s & ", " & prega & ", '" & Trim(Left(d4!sku, 4)) & "'"          'jv082415
                        s = s & ", " & d4(2) & ", " & pgrp & ", " & d4(2) - pgrp & ")"
                    End If
                    Sdb.Execute s               'jv060916
                End If
            End If
            d2.Close
            d4.MoveNext
        Loop
    End If
    d4.Close
    
    'Ante Room --------------------------------------------------------------------
    s = "select sku,count(*) from rackpos where sku > '000'"
    s = s & " and bbc = 'Y'"        'jv111611
    s = s & " and rackno in (select id from racks where aisle = 'M' and rack = 'ANTE')"
    s = s & " group by sku having count(*) > 0"
    Set d4 = Wdb.Execute(s)                 'jv060916
    If d4.BOF = False Then
        d4.MoveFirst
        Do Until d4.EOF
            s = "select * from skumast where sku = '" & Trim(Left(d4!sku, 4)) & "'"         'jv082415
            Set d2 = Sdb.Execute(s)         'jv060916
            If d2.BOF = True Then
                MsgBox "What is this sku? " & d4!sku, vbOKOnly, "Garbage found in ANTE room."
            Else
                s = "select product,count(*) from paltasks where area = 'FORKLIFT'"
                s = s & " and status = 'PEND' and target = 'STAGING'"
                s = s & " and product >= '" & d4!sku & "'"
                s = s & " and product < '" & d4!sku & "ZZZ'"
                s = s & " and source = 'ANTE ROOM'"
                s = s & " group by product having count(*) > 0"
                Set d3 = Wdb.Execute(s)     'jv060916
                If d3.BOF = False Then
                    d3.MoveFirst
                    pgrp = d3(1)
                Else
                    pgrp = 0
                End If
                d3.Close
                s = "select * from whstotals where sku = '" & Trim(Left(d4!sku, 4)) & "' and whs_num = " & pante        'jv082415
                Set d1 = Sdb.Execute(s)         'jv060916
                If d1.BOF = True Then
                    pkey = wd_seq("whstotals", Form1.shipdb)
                    s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                    s = s & ", " & pante & ", '" & Trim(Left(d4!sku, 4)) & "'"              'jv082415
                    s = s & ", " & d4(1) & ", " & pgrp & ", " & d4(1) - pgrp & ")"
                Else
                    s = "Update whstotals set count_qty = count_qty + " & d4(1)
                    s = s & ", avail = avail + " & d4(1) & " Where id = " & d1!id
                End If
                d1.Close
                Sdb.Execute s               'jv060916
            End If
            d2.Close
            d4.MoveNext
        Loop
    End If
    d4.Close
    
    'Snack Plant ------------------------------------------------------------------------
    s = "select sku,count(*) from rackpos where sku > '000'"
    s = s & " and bbc = 'Y'"                'jv111611
    s = s & " and rackno in (select id from racks where aisle = 'M' and rack = 'SP')"
    s = s & " group by sku having count(*) > 0"
    Set d4 = Wdb.Execute(s)                 'jv060916
    If d4.BOF = False Then
        d4.MoveFirst
        Do Until d4.EOF
            s = "select * from skumast where sku = '" & Trim(Left(d4!sku, 4)) & "'"         'jv082415
            Set d2 = Sdb.Execute(s)         'jv060916
            If d2.BOF = True Then
                MsgBox "What is this sku? " & d4!sku, vbOKOnly, "Garbage found at Snack Plant."
            Else
                If d2!whs_num = psp Then
                    s = "select product,count(*) from paltasks where area = 'FORKLIFT'"
                    s = s & " and status = 'PEND' and target = 'STAGING'"
                    s = s & " and product >= '" & d4!sku & "'"
                    s = s & " and product < '" & d4!sku & "ZZZ'"
                    s = s & " and source = 'SNACK PLANT'"
                    s = s & " group by product having count(*) > 0"
                    Set d3 = Wdb.Execute(s) 'jv060916
                    If d3.BOF = False Then
                        d3.MoveFirst
                        pgrp = d3(1)
                    Else
                        pgrp = 0
                    End If
                    d3.Close
                    s = "select * from whstotals where sku = '" & Trim(Left(d4!sku, 4)) & "' and whs_num = " & psp      'jv082415
                    Set d1 = Sdb.Execute(s)     'jv060916
                    If d1.BOF = True Then
                        pkey = wd_seq("whstotals", Form1.shipdb)
                        s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                        s = s & ", " & psp & ", '" & Trim(Left(d4!sku, 4)) & "'"            'jv082415
                        s = s & ", " & d4(1) & ", " & pgrp & ", " & d4(1) - pgrp & ")"
                    Else
                        s = "Update whstotals set count_qty = count_qty + " & d4(1)
                        s = s & ", avail = avail + " & d4(1) & " Where id = " & d1!id
                    End If
                    d1.Close
                    Sdb.Execute s           'jv060916
                End If
            End If
            d2.Close
            d4.MoveNext
        Loop
    End If
    d4.Close
    
    '4 Ways
    s = "select sku,sum(qty4) from racks where hold=0 and sku > '000'"
    s = s & " group by sku having sum(qty4) > 0"
    Set d4 = Wdb.Execute(s)                 'jv060916
    If d4.BOF = False Then
        d4.MoveFirst
        Do Until d4.EOF
            s = "select * from skumast where sku = '" & Trim(Left(d4!sku, 4)) & "'"         'jv082415
            Set d2 = Sdb.Execute(s)         'jv060916
            If d2.BOF = True Then
                MsgBox "Bad sku: " & d4!sku, vbOKOnly, "Garbage in BRW racks.."
            Else
                s = "select sku,sum(order_qty - ship_plt_qty) from ship_rack"
                s = s & " where sku = '" & d2!sku & "' and ship_status"
                s = s & " not in ('DONE','CANC') and bbp = 'N' group by sku"
                s = s & " having sum(order_qty - ship_plt_qty) > 0"
                Set d3 = Wdb.Execute(s)     'jv060916
                If d3.BOF = True Then
                    pgrp = 0
                Else
                    pgrp = d3(1)
                End If
                d3.Close
                pkey = wd_seq("whstotals", Form1.shipdb)
                s = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                s = s & ", " & p4way & ", '" & Trim(Left(d4!sku, 4)) & "'"                  'jv082415
                s = s & ", " & d4(1) & ", " & pgrp & ", " & d4(1) - pgrp & ")"
                Sdb.Execute s               'jv060916
            End If
            d2.Close
            d4.MoveNext
        Loop
    End If
    d4.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "import_bc_racks", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " import_bc_racks - Error Number: " & eno
        End
    End If
End Sub

Private Sub badirtrl_sql(runlist As String)
    Dim tds As adodb.Recordset, tds2 As adodb.Recordset
    Dim kdb As adodb.Connection, s As String, z As Long
    Dim tno As String, dbr As Integer, tdate As String, bagc As String
    Set kdb = CreateObject("ADODB.Connection")
    kdb.Open Form1.baship
    
    'Process Trailers
    s = "delete from trailers where runid in " & runlist
    kdb.Execute s
    s = "select * from trailers where runid in " & runlist
    Set tds = Sdb.Execute(s)        'jv060916
    If tds.BOF = False Then
        tds.MoveFirst
        Do Until tds.EOF
            tno = tds!trlno
            dbr = tds!branch
            tdate = Format(tds!shipdate, "mm-dd-yyyy")
            bagc = "T" & mid(tdate, 4, 2) & Format(dbr, "00") & mid(tno, 2, 1)
            z = wd_seq("Trailers", Form1.baship)
            s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno"
            s = s & ", sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & z
            s = s & ", " & tds!runid
            s = s & ", '" & bagc & "'"
            s = s & ", " & tds!plant
            s = s & ", " & tds!branch
            s = s & ", '" & tds!account & "'"
            s = s & ", '" & tds!shipdate & "'"
            s = s & ", '" & tds!trlno & "'"
            s = s & ", '" & tds!sku & "'"
            s = s & ", " & tds!pallets
            s = s & ", " & tds!wraps
            s = s & ", " & tds!units
            s = s & ", " & tds!whs_num
            s = s & ", 'N', 'N')"
            kdb.Execute s
            tds.MoveNext
        Loop
    End If
    tds.Close
    
    'Process Branch Orders
    s = "select * from runs where id in " & runlist & " and loaded = '51'"
    Set tds = Sdb.Execute(s)        'jv060916
    If tds.BOF = False Then
        tds.MoveFirst
        Do Until tds.EOF
            s = "delete from brorders where plant = " & tds!loaded
            s = s & " and branch = " & tds!Destination
            s = s & " and orddate = '" & Format(tds!trldate, "m-d-yyyy") & "'"
            kdb.Execute s
            
            s = "select * from brorders where plant = " & tds!loaded
            s = s & " and branch = " & tds!Destination
            s = s & " and orddate = '" & Format(tds!trldate, "m-d-yyyy") & "'"
            Set tds2 = Sdb.Execute(s)           'jv060916
            If tds2.BOF = False Then
                tds2.MoveFirst
                Do Until tds2.EOF
                    z = wd_seq("Brorders", Form1.baship)
                    s = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty"
                    s = s & ", grpqty, netqty, altflag, partqty) Values (" & z
                    s = s & ", " & tds2!plant
                    s = s & ", " & tds2!branch
                    s = s & ", '" & tds2!account & "'"
                    s = s & ", '" & tds2!sku & "'"
                    s = s & ", '" & tds2!orddate & "'"
                    s = s & ", " & tds2!ordqty
                    s = s & ", " & tds2!grpqty
                    s = s & ", " & tds2!netqty
                    s = s & ", '" & tds2!altflag & "'"
                    s = s & ", " & tds2!partqty & ")"
                    kdb.Execute s
                    tds2.MoveNext
                Loop
            End If
            tds2.Close
            tds.MoveNext
        Loop
    End If
    tds.Close
    kdb.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "badirtrl_sql", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " badirtrl_sql - Error Number: " & eno
        End
    End If
End Sub

Private Sub sydirtrl_sql(runlist As String)
    Dim tds As adodb.Recordset, tds2 As adodb.Recordset
    Dim adb As adodb.Connection, s As String, z As Long
    Dim tno As String, dbr As Integer, tdate As String, bagc As String
    Set adb = CreateObject("ADODB.Connection")
    adb.Open Form1.syship
    
    'Process Trailers
    s = "delete from trailers where runid in " & runlist
    adb.Execute s
    s = "select * from trailers where runid in " & runlist
    Set tds = Sdb.Execute(s)        'jv060916
    If tds.BOF = False Then
        tds.MoveFirst
        Do Until tds.EOF
            tno = tds!trlno
            dbr = tds!branch
            tdate = Format(tds!shipdate, "mm-dd-yyyy")
            bagc = "T" & mid(tdate, 4, 2) & Format(dbr, "00") & mid(tno, 2, 1)
            z = wd_seq("Trailers", Form1.syship)
            s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno"
            s = s & ", sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & z
            s = s & ", " & tds!runid
            s = s & ", '" & bagc & "'"
            s = s & ", " & tds!plant
            s = s & ", " & tds!branch
            s = s & ", '" & tds!account & "'"
            s = s & ", '" & tds!shipdate & "'"
            s = s & ", '" & tds!trlno & "'"
            s = s & ", '" & tds!sku & "'"
            s = s & ", " & tds!pallets
            s = s & ", " & tds!wraps
            s = s & ", " & tds!units
            s = s & ", " & tds!whs_num
            s = s & ", 'N', 'N')"
            adb.Execute s
            tds.MoveNext
        Loop
    End If
    tds.Close
    
    'Process Branch Orders
    s = "select * from runs where id in " & runlist & " and loaded = '52'"
    Set tds = Sdb.Execute(s)                'jv060916
    If tds.BOF = False Then
        tds.MoveFirst
        Do Until tds.EOF
            s = "delete from brorders where plant = " & tds!loaded
            s = s & " and branch = " & tds!Destination
            s = s & " and orddate = '" & Format(tds!trldate, "m-d-yyyy") & "'"
            adb.Execute s
            
            s = "select * from brorders where plant = " & tds!loaded
            s = s & " and branch = " & tds!Destination
            s = s & " and orddate = '" & Format(tds!trldate, "m-d-yyyy") & "'"
            Set tds2 = Sdb.Execute(s)       'jv060916
            If tds2.BOF = False Then
                tds2.MoveFirst
                Do Until tds2.EOF
                    z = wd_seq("Brorders", Form1.syship)
                    s = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty"
                    s = s & ", grpqty, netqty, altflag, partqty) Values (" & z
                    s = s & ", " & tds2!plant
                    s = s & ", " & tds2!branch
                    s = s & ", '" & tds2!account & "'"
                    s = s & ", '" & tds2!sku & "'"
                    s = s & ", '" & tds2!orddate & "'"
                    s = s & ", " & tds2!ordqty
                    s = s & ", " & tds2!grpqty
                    s = s & ", " & tds2!netqty
                    s = s & ", '" & tds2!altflag & "'"
                    s = s & ", " & tds2!partqty & ")"
                    adb.Execute s
                    tds2.MoveNext
                Loop
            End If
            tds2.Close
            tds.MoveNext
        Loop
    End If
    tds.Close
    adb.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sydirtrl_sql", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sydirtrl_sql - Error Number: " & eno
        End
    End If
End Sub

Private Sub form_memo(memx As String)
    Dim i As Long, k As Long, filx As String
    List1.Clear: List2.Clear
    i = 1: k = 1
    If Len(memx) < 72 Then
        List2.AddItem memx
        Exit Sub
    End If
    Do Until i = 0
        i = InStr(i, memx, " ", vbBinaryCompare)
        If i = 0 Then Exit Do
        List1.AddItem Trim(mid(memx, k, i - k))
        k = i
        i = i + 1
    Loop
    List1.AddItem Trim(mid(memx, k, Len(memx) - k + 1))
    filx = ""
    For i = 0 To List1.ListCount - 1
        If Len(filx & List1.List(i)) > 72 Then
            List2.AddItem filx
            filx = List1.List(i) & " "
        Else
            filx = filx & List1.List(i) & " "
        End If
    Next i
    List2.AddItem filx
End Sub

Private Sub addtrl_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim mplant As String, mbranch As String, msku As String, mtrl As String, pkey As Long
    Dim mdate As String, mrun As Long, munits As Integer, mpals As String, bagc As String
    Dim mwhs As String                                              'jv121015
    On Error GoTo vberror
    mdate = InputBox("Trailer Date:", "Trailer Date", Format(Now, "m-d-yyyy"))
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid Date Format Used!", vbOKOnly + vbExclamation, "Sorry, try again..."
        Exit Sub
    End If
    mplant = InputBox("Plant Code:", "Plant Code", Form1.plantno)
    If Len(mplant) = 0 Then
        Exit Sub
    End If
    Set ds = Sdb.Execute("select * from plants where plant = " & mplant)    'jv060916
    If ds.BOF = True Then
        ds.Close
        MsgBox "Invalid Plant Code: " & mplant, vbOKOnly + vbExclamation, "Sorry, try again..."
        Exit Sub
    End If
    ds.Close
    mbranch = InputBox("Branch Code:", "Branch Code", "04")
    If Len(mbranch) = 0 Then
        Exit Sub
    End If
    Set ds = Sdb.Execute("select * from branches where branch = " & mbranch)    'jv060916
    If ds.BOF = True Then
        ds.Close
        MsgBox "Invalid Branch Code: " & mbranch, vbOKOnly + vbExclamation, "Sorry, try again..."
        Exit Sub
    End If
    ds.Close
    mtrl = InputBox("Trailer #:", "Trailer Number", "#1")
    If Len(mtrl) = 0 Then
        Exit Sub
    End If
    If Len(mtrl) > 2 Then
        MsgBox "Please Limit Trailer # to 2 characters.", vbOKOnly + vbInformation, "Sorry, try again..."
        Exit Sub
    End If
    sqlx = "select * from trailers where shipdate = '" & mdate & "'"
    sqlx = sqlx & " and plant = " & mplant
    sqlx = sqlx & " and branch = " & mbranch
    sqlx = sqlx & " and trlno = '" & mtrl & "'"
    Set ds = Sdb.Execute(sqlx)      'jv060916
    If ds.BOF = False Then
        ds.Close
        MsgBox "Trailer Already Exists For The Specified Date.", vbOKOnly + vbExclamation, "Sorry, Check trailers..."
        Exit Sub
    End If
    msku = InputBox("SKU:", "SKU Number", "777")
    If Len(msku) = 0 Then
        Exit Sub
    End If
    Set ds = Sdb.Execute("select * from skumast where sku = '" & msku & "'")        'jv060916
    If ds.BOF = True Then
        ds.Close
        MsgBox "Invalid SKU: " & msku, vbOKOnly + vbExclamation, "Sorry, try again..."
        Exit Sub
    Else
        ds.MoveFirst
        munits = ds!pallet
    End If
    ds.Close
    mpals = InputBox("Number of pallets:", "Pallets", "1")
    If Len(mpals) = 0 Then
        Exit Sub
    End If
    
    If mplant = "51" Then                                                   'jv121015
        mwhs = "14"                                                         'jv121015
    Else                                                                    'jv121015
        If mplant = "50" Then                                               'jv121015
            mwhs = InputBox("Warehouse:", "Warehouse Number....", "5")      'jv121015
        Else                                                                'jv121015
            mwhs = "0"                                                      'jv121015
        End If                                                              'jv121015
    End If                                                                  'jv121015
    If Len(mwhs) = 0 Then Exit Sub                                          'jv121015
    
    'R12
    'pkey = wd_seq("runs", Form1.shipdb)
    pkey = wd_seq("Oratkt", Form1.schdb)
    sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    sqlx = sqlx & " Values (" & pkey & ", " & mplant & ", " & mbranch
    sqlx = sqlx & ", 'Branch-" & mbranch & "'"
    sqlx = sqlx & ", '" & mtrl & "'"
    sqlx = sqlx & ", " & mpals
    sqlx = sqlx & ", '" & Format(mdate, "m-d-yyyy") & "'"
    sqlx = sqlx & ", '12:00 PM'"
    sqlx = sqlx & ", 'Added by plant " & Form1.plantno & "'"
    sqlx = sqlx & ", '*')"
    Sdb.Execute sqlx                'jv060916
    mrun = pkey

    pkey = wd_seq("trailers", Form1.shipdb)
    sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku, pallets"
    sqlx = sqlx & ", wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey
    sqlx = sqlx & ", " & mrun
    If mplant <> "50" Then
        bagc = "T" & mid(Format(mdate, "mmddyyyy"), 3, 2)
        bagc = bagc & Format(Val(mbranch), "00")
        bagc = bagc & Right(mtrl, 1)
        sqlx = sqlx & ", '" & bagc & "'"
    Else
        sqlx = sqlx & ", 'A" & Format(mbranch, "00") & "-" & mtrl & "'"                'jv121015
    End If
    sqlx = sqlx & ", " & mplant
    sqlx = sqlx & ", " & mbranch
    sqlx = sqlx & ", '......'"
    sqlx = sqlx & ", '" & Format(mdate, "m-d-yyyy") & "'"
    sqlx = sqlx & ", '" & mtrl & "'"
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", " & Val(mpals)
    sqlx = sqlx & ", 0"
    sqlx = sqlx & ", " & munits * Val(mpals)
    'If mplant = "51" Then
    '    sqlx = sqlx & ", 14"
    'Else
    '        sqlx = sqlx & ", 0"
    'End If
    sqlx = sqlx & ", " & mwhs                                   'jv121015
    sqlx = sqlx & ", 'N', 'N')"
    Sdb.Execute sqlx                'jv060916
    
    MsgBox "Use the edit trailer tab to complete the trailer just added: " & mrun, vbOKOnly + vbInformation, "Trailer Added.."
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, addtrl.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addtrl_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub bobtotrl_Click()
    jobtotrl.jbob = "bob"
    jobtotrl.Show
End Sub

Private Sub brgrpmen_Click()
    PostBrGrps.Show
End Sub

Private Sub brnotes_Click()
    Dim sdir As String, spath As String, sfile As String, oline As String, i As Integer
    Dim ds As adodb.Recordset, s As String, bname As String, nr As Boolean
    'On Error GoTo vberror
    hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 3
    Screen.MousePointer = 11
    spath = Form1.webdir & "\orders\notes.??"
    sdir = Dir$(spath)
    Do While sdir <> ""
        bname = "..."
        s = "Select branchname from branches where branch = " & Right(sdir, 2)
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            bname = ds(0)
        End If
        ds.Close
        sfile = Form1.webdir & "\orders\" & sdir
        nr = True
        Open sfile For Input As #1
        Do Until EOF(1)
            Line Input #1, oline
            oline = Trim(oline)
            If nr = True Then
                hgrid.AddItem bname & Chr(9) & Format(FileDateTime(sfile), "m-d-yyyy") & Chr(9) & oline
                nr = False
            Else
                hgrid.AddItem Chr(9) & Chr(9) & oline
            End If
        Loop
        Close #1
        sdir = Dir$
    Loop
    hgrid.FormatString = "<Branch|^Date|<Notes"
    hgrid.ColWidth(0) = 2200
    hgrid.ColWidth(1) = 1200
    hgrid.ColWidth(2) = 9000
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "brnotes_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " brnotes_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub brord1_Click()
    Dim sqlx As String, i As Integer, mdate As String, mbr As String
    Dim ds As adodb.Recordset, s As String, ss As adodb.Recordset
    Dim rt As String, rh As String, rf As String, c As Integer
    On Error GoTo vberror
    mbr = InputBox("Branch Code:", "Enter Branch Code", "3")
    If Len(mbr) = 0 Or Val(mbr) = 0 Then Exit Sub
    mdate = InputBox$("Shipping Date", "Shipping Date", Form1.cdate)
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid date format....", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Form1.cdate = Format(mdate, "m-d-yyyy")
    
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    s = "select brorders.plant,plantname,brorders.branch,branchname,"
    s = s & "sum(ordqty),sum(partqty) from brorders,plants,branches"
    s = s & " where orddate = '" & mdate & "'"
    s = s & " and brorders.branch = " & Val(mbr)
    s = s & " and plants.plant = brorders.plant"
    s = s & " and branches.branch = brorders.branch"
    s = s & " group by brorders.plant,plantname,brorders.branch,branchname"
    s = s & " order by brorders.plant,brorders.branch"
    Set ds = Sdb.Execute(s)             'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Grid1.Rows > 2 Then Grid1.AddItem " "
            s = ds(2) & Chr(9) & ds!branchname
            s = s & " - " & ds(1) & Chr(9) & Chr(9) & Chr(9) & ds(4) & " Pallets"
            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(5) & " Wraps"
            Grid1.AddItem s
            s = "select brorders.sku,fgunit,fgdesc,ordqty,partqty,altflag"
            s = s & " from brorders,skumast"
            s = s & " where orddate = '" & mdate & "'"
            s = s & " and plant = " & ds(0)
            s = s & " and branch = " & ds(2)
            s = s & " and skumast.sku = brorders.sku"
            s = s & " order by brorders.sku"
            Set ss = Sdb.Execute(s)         'jv060916
            If ss.BOF = False Then
                ss.MoveFirst
                c = 1
                Do Until ss.EOF
                    If c = 1 Then
                        s = ss(0) & Chr(9)
                        s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                        s = s & ss!ordqty & Chr(9)
                        If ss!altflag = "Y" Then s = s & "#"
                        s = s & Chr(9)
                        If ss!partqty > 0 Then s = s & ss!partqty & " Wraps"
                        s = s & Chr(9)
                        c = 2
                    Else
                        s = s & ss(0) & Chr(9)
                        s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                        s = s & ss!ordqty & Chr(9)
                        If ss!altflag = "Y" Then s = s & "#"
                        s = s & Chr(9)
                        If ss!partqty > 0 Then s = s & ss!partqty & " Wraps"
                        s = s & Chr(9)
                        c = 1
                        Grid1.AddItem s
                    End If
                    ss.MoveNext
                Loop
            End If
            ss.Close
            If c = 2 Then Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.FormatString = "^|<|>|^|>|^|<|>|^|>"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 400
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 3000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 400
    Grid1.ColWidth(9) = 1000
    Grid1.FillStyle = flexFillRepeat
    For c = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(c, 2) <= " " Then
            Grid1.Row = c: Grid1.RowSel = c
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Grid1.BackColorFixed
        End If
    Next c
    
    rt = "Branch Orders"
    rh = Format(mdate, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        htdc(0) = "Yellow": gndc(0) = hcolor.BackColor
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, brord1.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " brord1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub brordprt_Click()
    Dim sqlx As String, i As Integer, mdate As String
    Dim ds As adodb.Recordset, s As String, ss As adodb.Recordset
    Dim rt As String, rh As String, rf As String, c As Integer
    On Error GoTo vberror
    mdate = InputBox$("Shipping Date", "Shipping Date", Form1.cdate)
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid date format....", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Form1.cdate = Format(mdate, "m-d-yyyy")
    
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    s = "select brorders.plant,plantname,brorders.branch,branchname,"
    s = s & "sum(ordqty),sum(partqty) from brorders,plants,branches"
    s = s & " where orddate = '" & mdate & "'"
    s = s & " and plants.plant = brorders.plant"
    s = s & " and branches.branch = brorders.branch"
    s = s & " group by brorders.plant,plantname,brorders.branch,branchname"
    s = s & " order by brorders.branch,brorders.plant"
    Set ds = Sdb.Execute(s)             'jv060916
    'MsgBox s
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Grid1.Rows > 2 Then Grid1.AddItem " "
            s = ds(2) & Chr(9) & ds!branchname
            s = s & " - " & ds(1) & Chr(9) & Chr(9) & Chr(9) & ds(4) & " Pallets"
            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(5) & " Wraps"
            Grid1.AddItem s
            s = "select brorders.sku,fgunit,fgdesc,ordqty,partqty,altflag"
            s = s & " from brorders,skumast"
            s = s & " where orddate = '" & mdate & "'"
            s = s & " and plant = " & ds(0)
            s = s & " and branch = " & ds(2)
            s = s & " and skumast.sku = brorders.sku"
            s = s & " order by brorders.sku"
            Set ss = Sdb.Execute(s)         'jv060916
            If ss.BOF = False Then
                ss.MoveFirst
                c = 1
                Do Until ss.EOF
                    If c = 1 Then
                        s = ss(0) & Chr(9)
                        s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                        s = s & ss!ordqty & Chr(9)
                        If ss!altflag = "Y" Then s = s & "#"
                        s = s & Chr(9)
                        If ss!partqty > 0 Then s = s & ss!partqty & " Wraps"
                        s = s & Chr(9)
                        c = 2
                    Else
                        s = s & ss(0) & Chr(9)
                        s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                        s = s & ss!ordqty & Chr(9)
                        If ss!altflag = "Y" Then s = s & "#"
                        s = s & Chr(9)
                        If ss!partqty > 0 Then s = s & ss!partqty & " Wraps"
                        s = s & Chr(9)
                        c = 1
                        Grid1.AddItem s
                    End If
                    ss.MoveNext
                Loop
            End If
            ss.Close
            If c = 2 Then Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.FormatString = "^|<|>|^|>|^|<|>|^|>"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 400
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 3000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 400
    Grid1.ColWidth(9) = 1000
    Grid1.FillStyle = flexFillRepeat
    For c = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(c, 2) <= " " Then
            Grid1.Row = c: Grid1.RowSel = c
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Grid1.BackColorFixed
        End If
    Next c
    
    rt = "Branch Orders"
    rh = Format(mdate, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        htdc(0) = "Yellow": gndc(0) = hcolor.BackColor
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "brordprt_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " brordprt_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub bsolow_Click()
    check_sku_orders
End Sub

Private Sub cantldate_Click()
    Dim pdate As String
    On Error GoTo vberror
    pdate = InputBox("Please enter a valid date.", "Cancel Production Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format.", vbOKOnly + vbExclamation, "Sorry, try again.."
        Exit Sub
    End If
    Screen.MousePointer = 11
    Wdb.Execute "delete from prodrcv where proddate = '" & pdate & "'"      'jv060916
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, cantldate.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cantldate_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub clrbrords_Click()
    Dim sqlx As String, pdate As String
    On Error GoTo vberror
    pdate = InputBox("Order Date:", "Clear Orders", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format ...", vbOKOnly + vbInformation, "Sorry....."
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Screen.MousePointer = 11
    sqlx = "Delete From Brorders where orddate = '" & pdate & "'"
    Sdb.Execute sqlx
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, clrbrords.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " clrbrords_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub clrhplog_Click()
    If MsgBox("Ok to clear home page user log?", vbQuestion + vbYesNo, "Are you sure...") = vbNo Then Exit Sub
    Open Form1.webdir & "\userlog" For Output As #1
    Print #1, Format(Now, "mm-dd-yyyy hh:mm am/pm") & " - Log cleared......."
    Close #1
End Sub

Private Sub clrtrls_Click()
    Dim pdate As String, sqlx As String, ds As adodb.Recordset
    On Error GoTo vberror
    pdate = InputBox("Enter Trailer Date:", "Trailer Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format: " & pdate, vbOKOnly + vbExclamation, "Sorry, Try Again..."
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Screen.MousePointer = 11
    sqlx = "delete from trailers where plant = " & Val(Form1.plantno) & " and shipdate = '" & pdate & "'"
    sqlx = sqlx & " and Ra_flag = 'Y'"
    Sdb.Execute sqlx                    'jv060916
    
    If Val(Form1.plantno) = 50 Then
        sqlx = "delete from trailers where plant <> 50 and shipdate = '" & pdate & "'"
        Sdb.Execute sqlx                'jv060916
        sqlx = "delete from trailers where plant = 50 and shipdate = '" & pdate & "'"
        sqlx = sqlx & " and branch = 1"
        Sdb.Execute sqlx                'jv060916
    End If
    'Clean up branch orders  3-25-2003
    sqlx = "Delete from brorders where orddate < '" & Format(DateAdd("d", -3, pdate), "m-d-yyyy") & "'"
    Sdb.Execute sqlx                    'jv060916
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, clrtrls.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " clrtrls_click - Error Number: " & eno
        End
    End If
End Sub


Private Sub Command1_Click()
    'imp_sylacauga_lowstock
    'postr12
    'check_sku_orders
    'truckwos.Show
    Dim q As String
    'q = "select * from runs"
    q = "select plant, branch, sum(pallets) as totpal from trailers group by plant,branch"
    'Call xmlfromado(Me.shipdb, "u:\xwork.xml", q)
    shipusermenu.Show
    'wdvalists.Show
End Sub

Private Sub coneords_Click()
    wdcones.Show
End Sub

Private Sub disxprod_Click()
    shipdisc.Show
End Sub

Private Sub drplist_Click()
    Dim sqlx As String, ds As adodb.Recordset, bno As adodb.Recordset
    Dim x, i As Integer, pdate As String
    Dim br As adodb.Recordset, ws As adodb.Recordset, bo As adodb.Recordset, pl As String
    On Error GoTo vberror
    pdate = InputBox("Please enter order date.", "Order Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date..", vbOKOnly + vbExclamation, "Aborting Request"
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Screen.MousePointer = 11
    Open Form1.tempdir & "\droppal.txt" For Output As #1
    Print #1, "Drops and ForkLift Pallets for orders: " & Format(pdate, "m-dd-yyyy")
    Print #1, " "
    sqlx = "Select * from warehouses where plant = 50 "
    sqlx = sqlx & " and whs in "
    sqlx = sqlx & "('REG','REGA','SP','DROP','SDRP')"
    Set ws = Sdb.Execute(sqlx)          'jv060916
    ws.MoveFirst
    Do Until ws.EOF
        pl = pl & Trim(ws!whs) & Space(6 - Len(Trim(ws!whs)))
        ws.MoveNext
    Loop
    Print #1, Space(22) & pl
    sqlx = "Select * from branches"
    sqlx = sqlx & " where branch in (select branch from brorders"
    sqlx = sqlx & " where plant = 50 and orddate = '" & pdate & "')"
    sqlx = sqlx & " order by branch"
    Set br = Sdb.Execute(sqlx)          'jv060916
    If br.BOF = False Then
        br.MoveFirst
        Do Until br.EOF
            pl = br!branchname & Space(18 - Len(br!branchname))
            ws.MoveFirst
            Do Until ws.EOF
                sqlx = "select branch,sum(netqty) from brorders"
                sqlx = sqlx & " where orddate = '" & pdate & "'"
                sqlx = sqlx & " and brorders.plant = 50"
                sqlx = sqlx & " and brorders.branch = " & br!branch
                sqlx = sqlx & " and netqty > 0"
                sqlx = sqlx & " and sku in "
                sqlx = sqlx & "(select sku from whstotals where whs_num = " & ws!whs_num & ")"
                sqlx = sqlx & " group by branch"
                Set bo = Sdb.Execute(sqlx)          'jv060916
                If bo.BOF = False Then
                    If bo(1) < 10 Then pl = pl & " "
                    If bo(1) < 100 Then pl = pl & " "
                    pl = pl & Space(3) & bo(1)
                Else
                    pl = pl & Space(5) & "."
                End If
                bo.Close
                ws.MoveNext
            Loop
            Print #1, pl
            br.MoveNext
        Loop
    End If
    br.Close: ws.Close
    Close #1
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\droppal.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "drplist_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " drplist_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edbillc_Click()
    Call EdBills.refresh_grid1(Format(Now, "m-d-yyyy"))
    EdBills.Show
End Sub

Private Sub edbranch_Click()
    branchconf.Show
End Sub

Private Sub edbrorders_Click()
    Brorders.Show
End Sub

Private Sub edbrprod_Click()
    brprods.Show
End Sub

Private Sub edgroups_Click()
    Editgroups.Show
End Sub

Private Sub edjobbing_Click()
    edjob.Show
End Sub

Private Sub edop_Click()
    oplist.Show
End Sub

Private Sub edplants_Click()
    Plants.Show
End Sub

Private Sub edplbranch_Click()
    Plantbranch.Show
End Sub

Private Sub edplskus_Click()
    Plantskus.Show
End Sub

Private Sub edprsource_Click()
    Prodsources.Show
End Sub

Private Sub edsku_Click()
    skulist.Show
End Sub

Private Sub edtrls_Click()
    Edittrl.Show
End Sub

Private Sub edtrlsht_Click()
    Trsheet.Show
End Sub

Private Sub edusers_Click()
    Shipuser.Show
End Sub

Private Sub edwhs_Click()
    Warehouses.Show
End Sub

Private Sub edwhstotals_Click()
    Whstotals.Show
End Sub

Private Sub eopwks_Click()
    'Form2.Show
End Sub

Private Sub Form_Load()
    Dim f As String, urole As String
    Dim ret As Long, s As String, i As Integer
    Dim lpbuff As String * 25
    localAppDataPath = Environ("LOCALAPPDATA") & "\Shipping"
    htmlTempFile = localAppDataPath & "\htmltemp.htm"
    check_hax
    ret = GetUserName(lpbuff, 25)
    Me.userid = Left(lpbuff, InStr(lpbuff, Chr(0)) - 1)
    'Me.userid = "BAUSER"
    Label10 = ".... "
    Form1.cdate = Format(Now, "m-d-yyyy")
    Form1.cgrp = "..."
    If Me.userid = "jvierus" Or Me.userid = "rlhalfmann" Then
        Command1.Visible = True
    Else
        Command1.Visible = False
    End If
    If UCase(Command()) = "BAUSER" Then
        Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
    ElseIf UCase(Command()) = "SYUSER" Then
        Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
    Else
        Dim site As String
        site = UCase(Left(Environ$("computername"), 2))
        'site = "BA"
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
    
    Line Input #1, f
    Do Until EOF(1)
        Line Input #1, f
        'f = LCase(f): f = Trim(f)
        If LCase(Left$(f, 6)) = "plant=" Then Form1.Caption = Form1.Caption & " " & Right(f, Len(f) - 6)
        If LCase(Left$(f, 10)) = "tempfiles=" Then tempdir = Right$(f, Len(f) - 10)
        If LCase(Left$(f, 8)) = "reports=" Then repdir = Right$(f, Len(f) - 8)
        If LCase(Left$(f, 7)) = "shipdb=" Then shipdb = Right$(f, Len(f) - 7)
        If LCase(Left$(f, 6)) = "schdb=" Then schdb = Right$(f, Len(f) - 6)
        If LCase(Left$(f, 5)) = "bbsr=" Then bbsr = Right$(f, Len(f) - 5)
        If LCase(Left$(f, 4)) = "ftp=" Then ftpdir = Right$(f, Len(f) - 4)
        If LCase(Left$(f, 7)) = "webdir=" Then webdir = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "trltrk=" Then trltrk = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "drvdir=" Then drvdir = Right(f, Len(f) - 7)
        If LCase(Left$(f, 8)) = "plantno=" Then plantno = Right(f, Len(f) - 8)
        If LCase(Left$(f, 7)) = "ratrls=" Then ratrls = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "baship=" Then baship = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "babbsr=" Then babbsr = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "syship=" Then syship = Right(f, Len(f) - 7)
        If LCase(Left$(f, 7)) = "sybbsr=" Then sybbsr = Right(f, Len(f) - 7)
        If LCase(Left$(f, 8)) = "pallogs=" Then pallogs = Right(f, Len(f) - 8)
        If LCase(Left$(f, 7)) = "srserv=" Then srserv = Right(f, Len(f) - 7)
    Loop
    Close #1
    labfmtfile = "\\BBC-03-FILESVR\SharedGroups\wd\bin\labfmt.txt"
    'Me.schdb = "ODBC;DATABASE=WDTruck;uid=bbctruck500;pwd=brenham500;DSN=truckwo"
    'Form1.schdb = "ODBC;DATABASE=WDTruck;DSN=wdtruck"
    'Form1.shipdb = "ODBC;DATABASE=WDship;DSN=wdship"
    'Form1.pallogs = "v:\testlogs\"
    'Form1.shipdb = "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    'Form1.shipdb = "ODBC;DATABASE=BAship;UID=bbcship501;PWD=Barrow501;DSN=wdship501"
    'Form1.shipdb = "ODBC;DATABASE=SYship;UID=bbcship502;PWD=Alabama502;DSN=wdship502"
    'Form1.baship = "ODBC;DATABASE=WDship;DSN=wdship"
    'Form1.syship = "ODBC;DATABASE=WDship;DSN=wdship"
    'Form1.syship = "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    'Form1.babbsr = "ODBC;DATABASE=BARacks;UID=bbcwd501;PWD=barrow501;DSN=wdsql501"
    'Form1.bbsr = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    'If Form1.plantno = 50 Then Form1.bbsr = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    'if Form1.plantno = 51 Then Form1.bbsr = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    
    Set Wdb = CreateObject("ADODB.Connection")
    Wdb.Open Me.bbsr
    wduserid = Me.userid
    Set Sdb = CreateObject("ADODB.Connection")
    Sdb.Open Me.shipdb
    
    ' Build local directory
    If DirExists(localAppDataPath) <> True Then
        MkDir (localAppDataPath)
    End If
    
    FrmGrid.FormatString = "^Form|^Top|^Left|^Height|^Width"
    FrmGrid.ColWidth(0) = 1000
    FrmGrid.ColWidth(1) = 800: FrmGrid.ColWidth(2) = 800
    FrmGrid.ColWidth(3) = 800: FrmGrid.ColWidth(4) = 800
    FrmGrid.Rows = 1
    On Error Resume Next
    Open localAppDataPath & "\shpforms.ini" For Input As #1
    If Err = 53 Then
        FrmGrid.AddItem "form1" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 2385 & Chr$(9) & 8595
        FrmGrid.AddItem "branches" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 7215 & Chr$(9) & 6855
        FrmGrid.AddItem "bravail" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 3600 & Chr$(9) & 4800
        FrmGrid.AddItem "brorders" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6945 & Chr$(9) & 7935
        FrmGrid.AddItem "brprods" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 4515 & Chr$(9) & 5655
        FrmGrid.AddItem "shipdisc" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 4515 & Chr$(9) & 5655
        FrmGrid.AddItem "editgroups" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 7515 & Chr$(9) & 10695
        FrmGrid.AddItem "edittrl" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6825 & Chr$(9) & 9105
        FrmGrid.AddItem "impords" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 4410 & Chr$(9) & 4800
        FrmGrid.AddItem "jobbing" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5415 & Chr$(9) & 8190
        FrmGrid.AddItem "jobtotrl" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6720 & Chr$(9) & 6435
        FrmGrid.AddItem "oplist" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6285 & Chr$(9) & 5775
        FrmGrid.AddItem "plantbranch" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5070 & Chr$(9) & 4590
        FrmGrid.AddItem "plants" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 2865 & Chr$(9) & 4860
        FrmGrid.AddItem "plantskus" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5805 & Chr$(9) & 7470
        FrmGrid.AddItem "postbrgrps" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6825 & Chr$(9) & 7260
        FrmGrid.AddItem "prodrcpts" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5940 & Chr$(9) & 8100
        FrmGrid.AddItem "prodsources" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 3930 & Chr$(9) & 4590
        FrmGrid.AddItem "rastrail" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 3420 & Chr$(9) & 8160
        FrmGrid.AddItem "skulist" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 7290 & Chr$(9) & 9105
        FrmGrid.AddItem "transched" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6060 & Chr$(9) & 8400
        FrmGrid.AddItem "trsheet" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6135 & Chr$(9) & 8880
        FrmGrid.AddItem "warehouses" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 4545 & Chr$(9) & 6285
        FrmGrid.AddItem "wdwan" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5910 & Chr$(9) & 5010
        FrmGrid.AddItem "whstotals" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 6855 & Chr$(9) & 7335
    Else
        Do Until EOF(1)
            Input #1, f, t, l, h, w
            FrmGrid.AddItem f & Chr$(9) & t & Chr$(9) & l & Chr$(9) & h & Chr$(9) & w
        Loop
    End If
    Close #1
    On Error GoTo 0
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "form1" Then
            Form1.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Form1.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Form1.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Form1.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    
    
    If UCase(Command()) = "BAUSER" Then
        Call menu_build("bauser")
    Else
        If UCase(Command()) = "SYUSER" Then
            Call menu_build("syluser")
        Else
            If UCase(Command()) = "ETONLY" Then
                Call menu_build("etonly")
            Else
                If UCase(Command()) = "PARTPALLETS" Then
                    Call menu_build("partpallets")
                Else
                    Call menu_build(Me.userid)
                    'Call menu_build("bvincik")
                End If
            End If
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
    
    'If check_version(Label13, Me.shipdb) = False Then
    '    s = "You are not using the most current version for " & Me.Caption & ".  You should use the"
    '    s = s & " update short-cut on your desktop to update this application.  Do you wish to "
    '    s = s & "continue with your current version?"
    '    If MsgBox(s, vbYesNo + vbQuestion, "Continue with current version...") = vbNo Then Unload Me
    '    'MsgBox "VErsion = true"
    'End If
    
    
    
    If Command() = "ETONLY" Then
        trailbill.Show
        Form1.Caption = Form1.Caption & " ETONLY"
        Form1.WindowState = 1
    End If
    If UCase(Command()) = "PARTPALLETS" Then
        partlabs.Show
        partlabs.WindowState = 2
        Form1.Caption = Form1.Caption & " PartPallets"
        Form1.WindowState = 1
    End If
    
End Sub

Private Sub Form_Resize()
    hgrid.Width = Me.Width - 120
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Form_Terminate()
    Call xitmenu_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call xitmenu_Click
End Sub

Private Sub gemmoh_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim ss As adodb.Recordset, bp As Long
    Dim query As String, i As Integer, k As Integer
    Dim userid As String, pwd As String, dsn As String
    Dim btot As Long, ptot As Long, t3gal As Long, ttray As Long
    Dim hf As String, rt As String, rh As String, rf As String, s As String
    On Error GoTo vberror
    hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 4
    hgrid.FillStyle = flexFillRepeat
    On Error Resume Next
    Open Form1.srserv & "\wd\bin\gemmodbc.ini" For Input As #1
    If Err = 53 Then
        MsgBox "Gemmodbc.ini File not found in wd\bin directory!", vbOKOnly + vbExclamation, "Request cancelled"
        Exit Sub
    Else
        Line Input #1, dsn
        Line Input #1, userid
        Line Input #1, pwd
        Close #1
    End If
    On Error GoTo 0
    'R12
    'dsn = "pbelle"
    'userid = "Apps"
    'pwd = "h0ly_c0w"
    'pwd = "papps"
    'MsgBox dsn & " " & userid & " " & pwd
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    Screen.MousePointer = 11
    sqlx = "select branch, branchname, fax, gemmsid from branches where gemmsid > '...' and fax >= '.'" 'jv102715
    Set ds = Sdb.Execute(sqlx)          'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid1.Cols = 4: Grid1.Clear
            hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 4
            'R12
            query = "select o.inventory_item_id, m.segment1, m.description, sum(o.transaction_quantity)"
            query = query & " from mtl_onhand_quantities o, mtl_system_items_b m, mtl_item_locations l"
            query = query & " where o.subinventory_code = '" & ds!gemmsid & "'"
            query = query & " and m.organization_id = o.organization_id"
            query = query & " and m.inventory_item_id = o.inventory_item_id"
            query = query & " and m.segment1 >= '100' and m.segment1 <= '999'"
            query = query & " and l.inventory_location_id = o.locator_id"                   'jv041515
            query = query & " and l.segment1 > 'FLOOR   '"                                  'jv041515
            query = query & " and l.segment1 < 'FLOORZZZ'"                                  'jv041515
            query = query & " group by o.inventory_item_id, m.segment1, m.description"
            query = query & " order by m.segment1"
            
            i = LoadGrid(Grid1, query, hstmt, 1, "")
            Open Form1.webdir & "\stock\goh." & Format(ds!branch, "00") For Output As #1
            hf = Form1.webdir & "\stock\goh" & Format(ds!branch, "00") & ".htm"
            'Open "s:\wd\test\goh." & Format(ds!branch, "00") For Output As #1
            'hf = "s:\wd\test\goh" & Format(ds!branch, "00") & ".htm"
            sqlx = "Oracle Inventory Report   "
            sqlx = sqlx & ds!branch & " " & ds!branchname & " " & Format(Now, "m-d-yyyy h:mm am/pm")
            rt = "Oracle Inventory Report"
            rf = ds!branch & " " & ds!branchname & "  Updated: " & Format(Now, "m-d-yyyy h:mm am/pm")
            Print #1, sqlx
            Print #1, " "
            btot = 0: ptot = 0: t3gal = 0: ttray = 0
            For k = 0 To Grid1.Rows - 1
                If Val(Grid1.TextMatrix(k, 1)) > 0 Then
                    btot = btot + Val(Grid1.TextMatrix(k, 3))
                End If
            Next k
            sqlx = "Total Units:  " & Format(btot, "##,###,##0")
            sqlx = sqlx & Space(53 - Len(sqlx))
            sqlx = sqlx & "Pallets   Units"
            Print #1, sqlx
            Print #1, " "
            For k = 0 To Grid1.Rows - 1
                If Val(Grid1.TextMatrix(k, 1)) > 0 Then
                    bp = 0
                    If Val(Grid1.TextMatrix(k, 3)) > 0 Then
                        sqlx = "select fgunit,fgdesc,pallet from skumast where sku = '" & Grid1.TextMatrix(k, 1) & "'"
                        Set ss = Sdb.Execute(sqlx)          'jv060916
                        If ss.BOF = False Then
                            ss.MoveFirst
                            If ss!pallet > 0 Then
                                If Left(ss!fgunit & ".", 1) = "3" Then
                                    t3gal = t3gal + Val(Grid1.TextMatrix(k, 3))
                                    bp = 0
                                Else
                                    If Left(ss!fgunit & ".", 1) = "T" Then
                                        ttray = ttray + Val(Grid1.TextMatrix(k, 3))
                                        bp = 0
                                    Else
                                        bp = Int((Val(Grid1.TextMatrix(k, 3)) / ss!pallet) + 0.999)
                                    End If
                                End If
                            End If
                            Grid1.TextMatrix(k, 2) = StrConv(ss!fgunit & " " & ss!fgdesc, vbProperCase)  'R12
                        End If
                        ss.Close
                    End If
                    sqlx = Grid1.TextMatrix(k, 1)
                    sqlx = sqlx & " " & Grid1.TextMatrix(k, 2)
                    sqlx = Left(sqlx, 50)
                    sqlx = sqlx & Space(50 - Len(sqlx))
                    sqlx = sqlx & Space(8 - Len(Format(bp, "#")))
                    sqlx = sqlx & Format(bp, "#")
                    Grid1.TextMatrix(k, 3) = Int(Val(Grid1.TextMatrix(k, 3)))
                    If Len(Grid1.TextMatrix(k, 3)) < 10 Then
                        sqlx = sqlx & Space(10 - Len(Grid1.TextMatrix(k, 3)))
                    End If
                    sqlx = sqlx & Grid1.TextMatrix(k, 3)
                    Print #1, sqlx
                    s = Grid1.TextMatrix(k, 1) & Chr(9)
                    s = s & Grid1.TextMatrix(k, 2) & Chr(9)
                    s = s & Format(bp, "#") & Chr(9)
                    s = s & Grid1.TextMatrix(k, 3)
                    hgrid.AddItem s
                    ptot = ptot + bp
                End If
            Next k
            Print #1, " "
            Print #1, "Total Units:  " & Format(btot, "##,###,##0")
            rh = "Total Units:  " & Format(btot, "##,###,##0") & "<BR>"
            Print #1, " "
            Print #1, "Pallet Space Summary ------------------------"
            rh = rh & "------ Pallet Space Summary ------<BR>"
            Print #1, "3 Gallons:   "; Tab(18); Space(6 - Len(Format(t3gal, "0"))); t3gal; " Units"; Tab(35); Space(6 - Len(Format(Int((t3gal / 60) + 0.999), "0"))); Int((t3gal / 60) + 0.999)
            rh = rh & "3 Gallons:   " & Format(t3gal, "0") & " Units   " & Int((t3gal / 60) + 0.999) & "; "
            Print #1, "Trays:       "; Tab(18); Space(6 - Len(Format(ttray, "0"))); ttray; " Units"; Tab(35); Space(6 - Len(Format(Int((ttray / 132) + 0.999), "0"))); Int((ttray / 132) + 0.999)
            rh = rh & "Trays:       " & Format(ttray, "0") & " Units   " & Int((ttray / 132) + 0.999) & "; "
            Print #1, "Other:       "; Tab(35); Space(6 - Len(Format(ptot, "0"))); ptot
            rh = rh & "Other:       " & ptot & "<BR>"
            ptot = ptot + Int((t3gal / 60) + 0.999)
            ptot = ptot + Int((ttray / 132) + 0.999)
            Print #1, "Total Pallets: "; Tab(35); Space(6 - Len(Format(ptot, "0"))); ptot
            rh = rh & "Total Pallets: " & ptot & "<BR>"
            Print #1, "Usable Capacity: "; Tab(35); Space(6 - Len(Format(Val(ds!fax), "0"))); Val(ds!fax)
            rh = rh & "Usable Capacity: " & Val(ds!fax) & "<BR>"
            If Len(ds!fax) > 0 And Val(ds!fax) > 0 Then
                Print #1, "Pct.           "; Tab(35); Space(7 - Len(Format(ptot / Val(ds!fax), ".000"))); Format(ptot / Val(ds!fax), ".000")
                rh = rh & "Pct.           " & Format(ptot / Val(ds!fax), ".000")
            Else
                Print #1, "Capacity Not Available"
                rh = rh & "Capacity Not Available"
            End If
            Close #1
            k = 0
            For i = 0 To hgrid.Rows - 1
                If Left(hgrid.TextMatrix(i, 3), 1) = "-" Then
                    hgrid.Row = i: hgrid.RowSel = i
                    hgrid.Col = 1: hgrid.ColSel = hgrid.Cols - 1
                    hgrid.CellBackColor = hcolor.BackColor
                    k = k + 1
                End If
            Next i
            If k = 1 Then rf = rf & "<BR>" & k & " item has negative quantity."
            If k > 1 Then rf = rf & "<BR>" & k & " items have negative quantities."
            hgrid.FormatString = "^SKU|<Product|^Pallets|^Units"
            hgrid.ColWidth(0) = 400
            hgrid.ColWidth(1) = 3000
            hgrid.ColWidth(2) = 1000
            hgrid.ColWidth(3) = 1000
            htdc(0) = "Yellow": gndc(0) = hcolor.BackColor
            Call htmlcolorgrid(Me, hf, hgrid, rt, rh, rf, "lemonchiffon", "linen", "white")
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "gemmoh_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " gemmoh_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub gemmsched_Click()
    Call schedlabels                'jv061914
End Sub

Private Sub grptotrl_Click()
    Dim tl As adodb.Recordset
    Dim gs As adodb.Recordset, ds2 As adodb.Recordset, ts As adodb.Recordset, ps As adodb.Recordset
    Dim mgroup As String, sqlx As String, parx As String, pdate As String
    Dim baflag As Boolean, batlist As String, pkey As Long
    Dim bplant As String, bbranch As String                                     'jv081916
    'On Error GoTo vberror
    baflag = False: syflag = False
    mgroup = UCase(InputBox$("Please enter groupcode", "Group Code"))
    'pdate = InputBox("Ship Date:", "Trailer Ship Date", Format(Now, "m-d-yyyy"))
    sqlx = "select * from trgroups where groupcode = '" & mgroup & "'"
    Set ts = Sdb.Execute(sqlx)          'jv060916
    sqlx = "select groupcode,groupitems.sku,qty1,whs1,qty2,whs2,qty3,whs3,qty4,whs4,pallet,numwrap"
    sqlx = sqlx & " from groupitems,skumast where groupcode = '" & mgroup & "'"
    sqlx = sqlx & " and groupitems.sku = skumast.sku"
    Set gs = Sdb.Execute(sqlx)          'jv060916
    If ts.BOF = True Or gs.BOF = True Then
        MsgBox "Groupcode not found....", vbOKOnly, "Invalid Group"
        gs.Close: ts.Close
        Exit Sub
    End If
    Screen.MousePointer = 11
    ts.MoveFirst
    sqlx = "select * from runs where id in (" & ts(1) & "," & ts(2) & "," & ts(3) & "," & ts(4) & ")"
    Set ds2 = Sdb.Execute(sqlx)         'jv060916
    If ds2.BOF = True Then
        ds2.Close
        gs.Close: ts.Close
        MsgBox "Runids in group not found in Trailer Schedule..", vbOKOnly, "Invalid Run"
        Exit Sub
    End If
    ds2.MoveFirst
    pdate = Format(ds2!trldate, "m-d-yyyy")
    ds2.Close
    
    sqlx = "Delete From Trailers Where Groupcode = '" & mgroup & "'"
    sqlx = sqlx & " and shipdate = '" & pdate & "'"
    Sdb.Execute sqlx            'jv060916
    sqlx = "select * from trailers where groupcode = '" & mgroup & "'"
    sqlx = sqlx & " and shipdate = '" & pdate & "'"
    Set tl = Sdb.Execute(sqlx)  'jv060916
    If tl.BOF = False Then
        gs.Close: ts.Close
        MsgBox "Detected Problems Clearing Previous Trailers..", vbOKOnly, "Try again.."
        Exit Sub
    End If
    
    gs.MoveFirst
    Do While Not gs.EOF
        If gs!qty1 > 0 And gs!whs1 > 0 Then
            Set ds2 = Sdb.Execute("select * from runs where id = " & ts!run1)       'jv060916
            ds2.MoveFirst
            If ds2!loaded = 50 Then bplant = "T10"                                  'jv081916
            If ds2!loaded = 51 Then bplant = "K10"                                  'jv081916
            If ds2!loaded = 52 Then bplant = "A10"                                  'jv081916
            If ds2!loaded = 51 Then baflag = True
            If ds2!loaded = 52 Then syflag = True
            If UCase(gs!sku) = "PAR" Then
                parx = "Select brorders.sku,partqty,numwrap from brorders,skumast"
                parx = parx & " Where branch = " & ds2("destination")
                parx = parx & " and orddate = '" & pdate & "'"
                parx = parx & " And partqty > 0"
                parx = parx & " And brorders.sku = skumast.sku"
                Set ps = Sdb.Execute(parx)              'jv060916
                If ps.BOF = False Then
                    ps.MoveFirst
                    Do Until ps.EOF
                        pkey = wd_seq("trailers", Form1.shipdb)
                        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                        sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                        sqlx = sqlx & " Values (" & pkey
                        sqlx = sqlx & ", " & ds2!id
                        sqlx = sqlx & ", '" & gs!groupcode & "'"
                        sqlx = sqlx & ", " & ds2!loaded
                        sqlx = sqlx & ", " & ds2!Destination
                        sqlx = sqlx & ", '......'"
                        sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                        sqlx = sqlx & ", '" & ds2!trlno & "'"
                        sqlx = sqlx & ", '" & UCase(ps!sku) & "'"
                        sqlx = sqlx & ", 0"
                        sqlx = sqlx & ", " & ps!partqty
                        sqlx = sqlx & ", " & ps!partqty * ps!numwrap
                        sqlx = sqlx & ", " & gs!whs1
                        sqlx = sqlx & ", 'N', 'N')"
                        Sdb.Execute sqlx                    'jv060916
                        sqlx = "Update bimp set onorder = onorder + "               'jv081916
                        sqlx = sqlx & Format(ps!partqty * ps!numwrap, "0")          'jv081916
                        sqlx = sqlx & " Where plantwhs = '" & bplant & "'"          'jv081916
                        sqlx = sqlx & " and branchwhs = '" & bbranch & "'"          'jv081916
                        sqlx = sqlx & " and sku = '" & ps!sku & "'"                 'jv081916
                        Sdb.Execute sqlx                                            'jv081916
                        ps.MoveNext
                    Loop
                End If
                ps.Close
            Else
                pkey = wd_seq("trailers", Form1.shipdb)
                sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                sqlx = sqlx & " Values (" & pkey
                sqlx = sqlx & ", " & ds2!id
                sqlx = sqlx & ", '" & gs!groupcode & "'"
                sqlx = sqlx & ", " & ds2!loaded
                sqlx = sqlx & ", " & ds2!Destination
                sqlx = sqlx & ", '......'"
                sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                sqlx = sqlx & ", '" & ds2!trlno & "'"
                sqlx = sqlx & ", '" & UCase(gs!sku) & "'"
                sqlx = sqlx & ", " & gs!qty1
                sqlx = sqlx & ", 0"
                sqlx = sqlx & ", " & gs!qty1 * gs!pallet
                sqlx = sqlx & ", " & gs!whs1
                sqlx = sqlx & ", 'N', 'N')"
                Sdb.Execute sqlx                'jv060916
                sqlx = "Update bimp set onorder = onorder + "                   'jv081916
                sqlx = sqlx & Format(gs!qty1 * gs!pallet, "0")                  'jv081916
                sqlx = sqlx & " Where plantwhs = '" & bplant & "'"              'jv081916
                sqlx = sqlx & " and branchwhs = '" & bbranch & "'"              'jv081916
                sqlx = sqlx & " and sku = '" & gs!sku & "'"                     'jv081916
                Sdb.Execute sqlx                                                'jv081916
            End If
            ds2.Close
        End If
        If gs!qty2 > 0 And gs!whs2 > 0 Then
            Set ds2 = Sdb.Execute("select * from runs where id = " & ts!run2)       'jv060916
            ds2.MoveFirst
            If ds2!loaded = 50 Then bplant = "T10"                                  'jv081916
            If ds2!loaded = 51 Then bplant = "K10"                                  'jv081916
            If ds2!loaded = 52 Then bplant = "A10"                                  'jv081916
            If ds2!loaded = 51 Then baflag = True
            If ds2!loaded = 52 Then syflag = True
            If UCase(gs!sku) = "PAR" Then
                parx = "Select brorders.sku,partqty,numwrap from brorders,skumast"
                parx = parx & " Where branch = " & ds2!Destination
                parx = parx & " and orddate = '" & pdate & "'"
                parx = parx & " And partqty > 0"
                parx = parx & " And brorders.sku = skumast.sku"
                Set ps = Sdb.Execute(parx)                      'jv060916
                If ps.BOF = False Then
                    ps.MoveFirst
                    Do Until ps.EOF
                        pkey = wd_seq("trailers", Form1.shipdb)
                        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                        sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                        sqlx = sqlx & " Values (" & pkey
                        sqlx = sqlx & ", " & ds2!id
                        sqlx = sqlx & ", '" & gs!groupcode & "'"
                        sqlx = sqlx & ", " & ds2!loaded
                        sqlx = sqlx & ", " & ds2!Destination
                        sqlx = sqlx & ", '......'"
                        sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                        sqlx = sqlx & ", '" & ds2!trlno & "'"
                        sqlx = sqlx & ", '" & UCase(ps!sku) & "'"
                        sqlx = sqlx & ", 0"
                        sqlx = sqlx & ", " & ps!partqty
                        sqlx = sqlx & ", " & ps!partqty * ps!numwrap
                        sqlx = sqlx & ", " & gs!whs2
                        sqlx = sqlx & ", 'N', 'N')"
                        Sdb.Execute sqlx                'jv060916
                        sqlx = "Update bimp set onorder = onorder + "               'jv081916
                        sqlx = sqlx & Format(ps!partqty * ps!numwrap, "0")          'jv081916
                        sqlx = sqlx & " Where plantwhs = '" & bplant & "'"          'jv081916
                        sqlx = sqlx & " and branchwhs = '" & bbranch & "'"          'jv081916
                        sqlx = sqlx & " and sku = '" & ps!sku & "'"                 'jv081916
                        Sdb.Execute sqlx                                            'jv081916
                        ps.MoveNext
                    Loop
                End If
                ps.Close
            Else
                pkey = wd_seq("trailers", Form1.shipdb)
                sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                sqlx = sqlx & " Values (" & pkey
                sqlx = sqlx & ", " & ds2!id
                sqlx = sqlx & ", '" & gs!groupcode & "'"
                sqlx = sqlx & ", " & ds2!loaded
                sqlx = sqlx & ", " & ds2!Destination
                sqlx = sqlx & ", '......'"
                sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                sqlx = sqlx & ", '" & ds2!trlno & "'"
                sqlx = sqlx & ", '" & UCase(gs!sku) & "'"
                sqlx = sqlx & ", " & gs!qty2
                sqlx = sqlx & ", 0"
                sqlx = sqlx & ", " & gs!qty2 * gs!pallet
                sqlx = sqlx & ", " & gs!whs2
                sqlx = sqlx & ", 'N', 'N')"
                Sdb.Execute sqlx                    'jv060916
                sqlx = "Update bimp set onorder = onorder + "                   'jv081916
                sqlx = sqlx & Format(gs!qty2 * gs!pallet, "0")                  'jv081916
                sqlx = sqlx & " Where plantwhs = '" & bplant & "'"              'jv081916
                sqlx = sqlx & " and branchwhs = '" & bbranch & "'"              'jv081916
                sqlx = sqlx & " and sku = '" & gs!sku & "'"                     'jv081916
                Sdb.Execute sqlx                                                'jv081916
            End If
            ds2.Close
        End If
        If gs!qty3 > 0 And gs!whs3 > 0 Then
            Set ds2 = Sdb.Execute("select * from runs where id = " & ts!run3)       'jv060916
            ds2.MoveFirst
            If ds2!loaded = 50 Then bplant = "T10"                                  'jv081916
            If ds2!loaded = 51 Then bplant = "K10"                                  'jv081916
            If ds2!loaded = 52 Then bplant = "A10"                                  'jv081916
            If ds2!loaded = 51 Then baflag = True
            If ds2!loaded = 52 Then syflag = True
            If UCase(gs!sku) = "PAR" Then
                parx = "Select brorders.sku,partqty,numwrap from brorders,skumast"
                parx = parx & " Where branch = " & ds2!Destination
                parx = parx & " and orddate = '" & pdate & "'"
                parx = parx & " And partqty > 0"
                parx = parx & " And brorders.sku = skumast.sku"
                Set ps = Sdb.Execute(parx)          'jv060916
                If ps.BOF = False Then
                    ps.MoveFirst
                    Do Until ps.EOF
                        pkey = wd_seq("trailers", Form1.shipdb)
                        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                        sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                        sqlx = sqlx & " Values (" & pkey
                        sqlx = sqlx & ", " & ds2!id
                        sqlx = sqlx & ", '" & gs!groupcode & "'"
                        sqlx = sqlx & ", " & ds2!loaded
                        sqlx = sqlx & ", " & ds2!Destination
                        sqlx = sqlx & ", '......'"
                        sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                        sqlx = sqlx & ", '" & ds2!trlno & "'"
                        sqlx = sqlx & ", '" & UCase(ps!sku) & "'"
                        sqlx = sqlx & ", 0"
                        sqlx = sqlx & ", " & ps!partqty
                        sqlx = sqlx & ", " & ps!partqty * ps!numwrap
                        sqlx = sqlx & ", " & gs!whs3
                        sqlx = sqlx & ", 'N', 'N')"
                        Sdb.Execute sqlx                    'jv060916
                        sqlx = "Update bimp set onorder = onorder + "               'jv081916
                        sqlx = sqlx & Format(ps!partqty * ps!numwrap, "0")          'jv081916
                        sqlx = sqlx & " Where plantwhs = '" & bplant & "'"          'jv081916
                        sqlx = sqlx & " and branchwhs = '" & bbranch & "'"          'jv081916
                        sqlx = sqlx & " and sku = '" & ps!sku & "'"                 'jv081916
                        Sdb.Execute sqlx                                            'jv081916
                        ps.MoveNext
                    Loop
                End If
                ps.Close
            Else
                pkey = wd_seq("trailers", Form1.shipdb)
                sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                sqlx = sqlx & " Values (" & pkey
                sqlx = sqlx & ", " & ds2!id
                sqlx = sqlx & ", '" & gs!groupcode & "'"
                sqlx = sqlx & ", " & ds2!loaded
                sqlx = sqlx & ", " & ds2!Destination
                sqlx = sqlx & ", '......'"
                sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                sqlx = sqlx & ", '" & ds2!trlno & "'"
                sqlx = sqlx & ", '" & UCase(gs!sku) & "'"
                sqlx = sqlx & ", " & gs!qty3
                sqlx = sqlx & ", 0"
                sqlx = sqlx & ", " & gs!qty3 * gs!pallet
                sqlx = sqlx & ", " & gs!whs3
                sqlx = sqlx & ", 'N', 'N')"
                Sdb.Execute sqlx                    'jv060916
                sqlx = "Update bimp set onorder = onorder + "                   'jv081916
                sqlx = sqlx & Format(gs!qty3 * gs!pallet, "0")                  'jv081916
                sqlx = sqlx & " Where plantwhs = '" & bplant & "'"              'jv081916
                sqlx = sqlx & " and branchwhs = '" & bbranch & "'"              'jv081916
                sqlx = sqlx & " and sku = '" & gs!sku & "'"                     'jv081916
                Sdb.Execute sqlx                                                'jv081916
            End If
            ds2.Close
        End If
        If gs!qty4 > 0 And gs!whs4 > 0 Then
            Set ds2 = Sdb.Execute("select * from runs where id = " & ts!run4)           'jv060916
            ds2.MoveFirst
            If ds2!loaded = 50 Then bplant = "T10"                                  'jv081916
            If ds2!loaded = 51 Then bplant = "K10"                                  'jv081916
            If ds2!loaded = 52 Then bplant = "A10"                                  'jv081916
            If ds2!loaded = 51 Then baflag = True
            If ds2!loaded = 52 Then syflag = True
            If UCase(gs!sku) = "PAR" Then
                parx = "Select brorders.sku,partqty,numwrap from brorders,skumast"
                parx = parx & " Where branch = " & ds2!Destination
                parx = parx & " and orddate = '" & pdate & "'"
                parx = parx & " And partqty > 0"
                parx = parx & " And brorders.sku = skumast.sku"
                Set ps = Sdb.Execute(parx)                      'jv060916
                If ps.BOF = False Then
                    ps.MoveFirst
                    Do Until ps.EOF
                        pkey = wd_seq("trailers", Form1.shipdb)
                        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                        sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                        sqlx = sqlx & " Values (" & pkey
                        sqlx = sqlx & ", " & ds2!id
                        sqlx = sqlx & ", '" & gs!groupcode & "'"
                        sqlx = sqlx & ", " & ds2!loaded
                        sqlx = sqlx & ", " & ds2!Destination
                        sqlx = sqlx & ", '......'"
                        sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                        sqlx = sqlx & ", '" & ds2!trlno & "'"
                        sqlx = sqlx & ", '" & UCase(ps!sku) & "'"
                        sqlx = sqlx & ", 0"
                        sqlx = sqlx & ", " & ps!partqty
                        sqlx = sqlx & ", " & ps!partqty * ps!numwrap
                        sqlx = sqlx & ", " & gs!whs4
                        sqlx = sqlx & ", 'N', 'N')"
                        Sdb.Execute sqlx                    'jv060916
                        sqlx = "Update bimp set onorder = onorder + "               'jv081916
                        sqlx = sqlx & Format(ps!partqty * ps!numwrap, "0")          'jv081916
                        sqlx = sqlx & " Where plantwhs = '" & bplant & "'"          'jv081916
                        sqlx = sqlx & " and branchwhs = '" & bbranch & "'"          'jv081916
                        sqlx = sqlx & " and sku = '" & ps!sku & "'"                 'jv081916
                        Sdb.Execute sqlx                                            'jv081916
                        ps.MoveNext
                    Loop
                End If
                ps.Close
            Else
                pkey = wd_seq("trailers", Form1.shipdb)
                sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate"
                sqlx = sqlx & ", trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag)"
                sqlx = sqlx & " Values (" & pkey
                sqlx = sqlx & ", " & ds2!id
                sqlx = sqlx & ", '" & gs!groupcode & "'"
                sqlx = sqlx & ", " & ds2!loaded
                sqlx = sqlx & ", " & ds2!Destination
                sqlx = sqlx & ", '......'"
                sqlx = sqlx & ", '" & Format(ds2!trldate, "m-d-yyyy") & "'"
                sqlx = sqlx & ", '" & ds2!trlno & "'"
                sqlx = sqlx & ", '" & UCase(gs!sku) & "'"
                sqlx = sqlx & ", " & gs!qty4
                sqlx = sqlx & ", 0"
                sqlx = sqlx & ", " & gs!qty4 * gs!pallet
                sqlx = sqlx & ", " & gs!whs4
                sqlx = sqlx & ", 'N', 'N')"
                Sdb.Execute sqlx                    'jv060916
                sqlx = "Update bimp set onorder = onorder + "                   'jv081916
                sqlx = sqlx & Format(gs!qty4 * gs!pallet, "0")                  'jv081916
                sqlx = sqlx & " Where plantwhs = '" & bplant & "'"              'jv081916
                sqlx = sqlx & " and branchwhs = '" & bbranch & "'"              'jv081916
                sqlx = sqlx & " and sku = '" & gs!sku & "'"                     'jv081916
                Sdb.Execute sqlx                                                'jv081916
            End If
            ds2.Close
        End If
        gs.MoveNext
    Loop
    batlist = "(" & ts!run1 & "," & ts!run2 & "," & ts!run3 & "," & ts!run4 & ")"
    tl.Close
    gs.Close: ts.Close ': db.Close
    If baflag = True Then
        Call badirtrl_sql(batlist)
    End If
    If syflag = True Then
        Call sydirtrl_sql(batlist)
    End If
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "grptotrl_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " grptotrl_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub homeupdt_Click()
    Bravail.Show
End Sub

Private Sub hpuserp_Click()
    Dim x
    FileCopy Form1.webdir & "\userlog", Form1.tempdir & "\userlog.txt"
    x = Shell("notepad.exe " & Form1.tempdir & "\userlog.txt", vbNormalFocus)
End Sub

Private Sub Image1_DblClick()
    Frame1.Visible = Not Frame1.Visible
End Sub

Private Sub impba_Click()
    Dim sqlx As String, inv As String, ds As adodb.Recordset, sdate As String
    Dim ba As adodb.Connection, br As adodb.Recordset, bs As adodb.Recordset, cq As Integer, gq As Integer
    Dim bawhs As Integer, pkey As Long
    On Error GoTo vberror
    Screen.MousePointer = 11
    sqlx = "select * from warehouses where plant = 51"
    Set ds = Sdb.Execute(sqlx)              'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        bawhs = ds!whs_num
    Else
        Screen.MousePointer = 0
        MsgBox "Cannot find a warehouse assigned to plant 51.", vbOKOnly, "Aborting..."
        ds.Close
        Exit Sub
    End If
    ds.Close
    sqlx = "delete from whstotals where whs_num = " & bawhs
    Sdb.Execute sqlx                        'jv060916
    
    ' Direct from BA Racks
    Set ba = CreateObject("ADODB.Connection")
    ba.Open Form1.babbsr
    sqlx = "select p.sku,count(*) from racks r, rackpos p where p.rackno = r.id"
    sqlx = sqlx & " and r.aisle <> 'M'"
    sqlx = sqlx & " and r.hold = 0 and p.sku > '0000' and p.sku < '9999'"               'jv082415
    sqlx = sqlx & " group by p.sku having count(*) > 0"
    Set br = ba.Execute(sqlx)
    If br.BOF = False Then
        br.MoveFirst
        Do Until br.EOF
            cq = br(1)
            sqlx = "select product,count(*) from paltasks where area = 'FORKLIFT'"
            sqlx = sqlx & " and status = 'PEND' and target = 'STAGING'"
            sqlx = sqlx & " and product >= '" & br(0) & "'"
            sqlx = sqlx & " and product < '" & br(0) & "ZZZ'"
            sqlx = sqlx & " and source not in ('ANTE ROOM', 'SNACK PLANT')"
            sqlx = sqlx & " group by product having count(*) > 0"
            Set bs = ba.Execute(sqlx)
            If bs.BOF = False Then
                bs.MoveFirst
                gq = bs(1)
            Else
                gq = 0
            End If
            bs.Close
            pkey = wd_seq("whstotals", Form1.shipdb)
            sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
            sqlx = sqlx & ", " & bawhs
            sqlx = sqlx & ", '" & br(0) & "'"
            sqlx = sqlx & ", " & cq
            sqlx = sqlx & ", " & gq
            sqlx = sqlx & ", " & cq - gq & ")"
            Sdb.Execute sqlx            'jv060916
            br.MoveNext
        Loop
    End If
    br.Close: ba.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "impba_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " impba_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub impbrorder_Click()
    Impords.Show
End Sub

Private Sub impreg_Click()
    Call import_bc_racks
End Sub

Private Sub impsched_Click()
    'Replaced with edit schedule - notes
End Sub

Private Sub impsr_Click()
    Call nt_availsr5
End Sub

Private Sub impsy_Click()
    Dim sqlx As String, inv As String, ds As adodb.Recordset, sdate As String
    Dim ba As adodb.Connection, br As adodb.Recordset, bs As adodb.Recordset, cq As Integer, gq As Integer
    Dim bawhs As Integer, cfile As String, s As String, i As Integer, psku As String, pkey As Long
    Dim sb As adodb.Connection, ss As adodb.Recordset
    Dim db5 As adodb.Connection, ds5 As adodb.Recordset
    Dim chp As Boolean
    'On Error GoTo vberror
    Screen.MousePointer = 11
    sqlx = "select * from warehouses where plant = 52"
    Set ds = Sdb.Execute(sqlx)          'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        bawhs = ds!whs_num
    Else
        Screen.MousePointer = 0
        MsgBox "Cannot find a warehouse assigned to plant 52.", vbOKOnly, "Aborting..."
        ds.Close
        Exit Sub
    End If
    ds.Close
    sqlx = "delete from whstotals where whs_num = " & bawhs
    Sdb.Execute sqlx            'jv060916
    
    ' Direct from SY Racks
    sqlx = "select * from whstotals where whs_num = " & bawhs
    Set ds = Sdb.Execute(sqlx)  'jv060916
    Set ba = CreateObject("ADODB.Connection")
    ba.Open Form1.sybbsr
    sqlx = "select p.sku,count(*) from racks r, rackpos p where p.rackno = r.id"
    sqlx = sqlx & " and r.aisle <> 'M'"
    sqlx = sqlx & " and r.hold = 0 and p.sku > '0000' and p.sku < '9999'"               'jv082415
    sqlx = sqlx & " group by p.sku having count(*) > 0"
    Set br = ba.Execute(sqlx)
    If br.BOF = False Then
        br.MoveFirst
        Do Until br.EOF
            cq = br(1)
            sqlx = "select product,count(*) from paltasks where area = 'FORKLIFT'"
            sqlx = sqlx & " and status = 'PEND' and target = 'STAGING'"
            sqlx = sqlx & " and product >= '" & br(0) & "'"
            sqlx = sqlx & " and product < '" & br(0) & "ZZZ'"
            sqlx = sqlx & " and source not in ('ANTE ROOM', 'SNACK PLANT')"
            sqlx = sqlx & " group by product having count(*) > 0"
            Set bs = ba.Execute(sqlx)
            If bs.BOF = False Then
                bs.MoveFirst
                gq = bs(1)
            Else
                gq = 0
            End If
            bs.Close
            pkey = wd_seq("whstotals", Form1.shipdb)
            sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
            sqlx = sqlx & ", " & bawhs
            sqlx = sqlx & ", '" & br(0) & "'"
            sqlx = sqlx & ", " & cq
            sqlx = sqlx & ", " & gq
            sqlx = sqlx & ", " & cq - gq & ")"
            Sdb.Execute sqlx                    'jv060916
            br.MoveNext
        Loop
    End If
    ds.Close
    br.Close
    ba.Close
    
    
    ' Use All Pallets.xls for CS5
    chp = False
    If MsgBox("Import CS5 pallets ON HOLD?", vbQuestion + vbYesNo, "Crane Pallets..") = vbYes Then chp = True
    Set db5 = CreateObject("ADODB.Connection")
    db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
    s = "SELECT tLocationData.sLocationID, "            'ds5(0)
    s = s & "tLaneData.iLevel,"                         'ds5(1)
    s = s & "tLaneData.iRow,"                           'ds5(2)
    s = s & "tLaneData.iBlock, "                        'ds5(3)
    s = s & "tContainerLocationData.iLocationID, "      'ds5(4)
    s = s & "tInventoryData.nQuantity, "                'ds5(5)
    s = s & "tLotData.dtProduction, "                   'ds5(6)
    s = s & "tItemMaster.sItemID,"                      'ds5(7)
    s = s & "tItemMaster.sItemDescription,"             'ds5(8)
    s = s & "tLaneLock.iLocked,"                        'ds5(9)
    s = s & "tLaneLock.sDescription,"                   'ds5(10)
    s = s & "count(*) "                                 'ds5(11)
    s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
    s = s & "tLotData, tItemMaster, tLaneLock"
    s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
    s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
    s = s & " AND tLaneLock.iLaneSysID = tLaneData.iLocationID"
    s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
    s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
    s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
    s = s & " GROUP BY tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    s = s & "tLotData.dtProduction, "
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLaneLock.iLocked, tLaneLock.sDescription"
    s = s & " ORDER BY tLocationData.sLocationID " ', tContainerLocationData.iPosition"
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF
            psku = Left(ds5(7), 3)                          'jv090215
            If Len(ds5(7)) > 3 Then                         'jv090215
                If mid(ds5(7), 4, 1) = "-" Then             'jv090215
                    psku = Left(ds5(7), 3)                  'jv090215
                End If                                      'jv090215
                If mid(ds5(7), 5, 1) = "-" Then             'jv090215
                    psku = Left(ds5(7), 4)                  'jv090215
                End If                                      'jv090215
            End If                                          'jv090215
            'psku = Trim(Left(ds5(7), 4))                                    'jv082415
            If chp = True Or ds5(9) = "0" Then
                s = "select * from whstotals where whs_num = " & bawhs
                s = s & " and sku = '" & psku & "'"
                Set ds = Sdb.Execute(s)                 'jv060916
                If ds.BOF = False Then
                    ds.MoveFirst
                    sqlx = "Update whstotals set count_qty = count_qty + " & ds5(11)
                    sqlx = sqlx & ", avail = avail + " & ds5(11) & " where id = " & ds!id
                Else
                    pkey = wd_seq("whstotals", Form1.shipdb)
                    sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail) Values (" & pkey
                    sqlx = sqlx & ", " & bawhs
                    sqlx = sqlx & ", '" & psku & "'"
                    sqlx = sqlx & ", " & ds5(11) & ", 0, " & ds5(11) & ")"
                End If
                ds.Close
                Sdb.Execute sqlx            'jv060916
            End If
            ds5.MoveNext
        Loop
    End If
    ds5.Close: db5.Close
    
    ' Add in-transit pallets to totals        jv 6-11-2008
    Call sylacauga_yard_countsheet
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 3)) > 0 Then
            s = "select * from whstotals where whs_num = " & bawhs
            s = s & " and sku = '" & Grid1.TextMatrix(i, 1) & "'"
            Set ds = Sdb.Execute(s)         'jv060916
            If ds.BOF = False Then
                sqlx = "Update whstotals set count_qty = count_qty + " & Val(Grid1.TextMatrix(i, 3))
                sqlx = sqlx & ", avail = avail + " & Val(Grid1.TextMatrix(i, 3))
                sqlx = sqlx & ", old_lot = " & Val(ds!old_lot & " ") + Val(Grid1.TextMatrix(i, 3))
                sqlx = sqlx & " where id = " & ds!id
            Else
                pkey = wd_seq("whstotals", Form1.shipdb)
                sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail, old_lot) Values (" & pkey
                sqlx = sqlx & ", " & bawhs
                sqlx = sqlx & ", '" & Grid1.TextMatrix(i, 1) & "'"
                sqlx = sqlx & ", " & Val(Grid1.TextMatrix(i, 3))
                sqlx = sqlx & ", 0"
                sqlx = sqlx & ", " & Val(Grid1.TextMatrix(i, 3))
                sqlx = sqlx & ", " & Val(Grid1.TextMatrix(i, 3)) & ")"
            End If
            ds.Close
            Sdb.Execute sqlx        'jv060916
        End If
    Next i
    
    ' Use Shipping Trailers for Order qtys    jv 11-28-2007
    If Right(Form1.syship, 4) = ".mdb" Then
        Set sb = OpenDatabase(Form1.syship)
    Else
        Set sb = CreateObject("ADODB.Connection")
        sb.Open (Form1.syship)
    End If
    s = "select * from whstotals where whs_num = " & bawhs
    Set ds = Sdb.Execute(s)         'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Right(Form1.syship, 4) = ".mdb" Then
                s = "select sku,sum(pallets) from trailers where plant = 52"
                s = s & " and sku = '" & ds!sku & "'"
                s = s & " and shipdate > '" & Format(Now, "m-d-yyyy") & "'"
                s = s & " group by sku having sum(pallets) > 0"
                'Set ss = sb.OpenRecordset(s)
            Else
                s = "select sku,sum(pallets) from trailers where plant = 52"
                s = s & " and sku = '" & ds!sku & "'"
                's = s & " and shipdate > '" & Format(Now, "m-d-yyyy") & "'"
                s = s & " and shipdate > '" & Format(Now, "m-d-yyyy") & "'"
                s = s & " group by sku having sum(pallets) > 0"
                'set ss = sb.Execute(s)
            End If
            Set ss = sb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                gq = ss(1)
            Else
                gq = 0
            End If
            ss.Close
            sqlx = "update whstotals set grp_qty = " & gq
            sqlx = sqlx & ", avail = " & ds!count_qty - gq
            sqlx = sqlx & " Where id = " & ds!id
            Sdb.Execute sqlx            'jv060916
            ds.MoveNext
        Loop
    End If
    sb.Close
    ds.Close
    '--------------------------------------
    
    imp_sylacauga_lowstock
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "impsy_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " impsy_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub instldate_Click()
    Dim pdate As String, psku As String, pdays As Integer
    Dim pqty As Long, psnk As Boolean, zid As Long
    Dim ds As adodb.Recordset, sqlx As String
    Dim ds2 As adodb.Recordset, plot As String
    On Error GoTo vberror
    pdate = InputBox("Please enter a valid date.", "Production Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Date entered as: " & pdate & " not recognized as valid.", vbOKOnly, "Sorry, try again..."
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    psku = InputBox("Please enter a valid sku.", "SKU for " & pdate, "777")
    If Len(psku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & psku & "'"
    Set ds = Sdb.Execute(sqlx)              'jv060916
    If ds.BOF = True Then
        MsgBox "SKU: " & psku & " not found in skumast.", vbOKOnly, "Sorry, cannot insert.."
        ds.Close
        Exit Sub
    End If
    sqlx = "select * from prodsources where source = " & ds!psource
    Set ds2 = Sdb.Execute(sqlx)             'jv060916
    If ds2.BOF = False Then
        pdays = ds2!days
        If ds2!tl_flag = "Y" Then                   'jv112015
            psnk = False
        Else
            If MsgBox("Is this product produced at the snack plant?", vbYesNo, ds!fgunit & " " & ds!fgdesc & " Snack Plant Product?") = vbYes Then
                psnk = True
            Else
                psnk = False
            End If
        End If
    Else
        pdays = Val(InputBox("How many days will this product be received.", ds!fgunit & " " & ds!fgdesc & " # Days", 1))
        If pdays < 1 Then
            ds2.Close: ds.Close ': db.Close
            Exit Sub
        End If
        If MsgBox("Is this product produced at the snack plant?", vbYesNo, ds!fgunit & " " & ds!fgdesc & " Snack Plant Product?") = vbYes Then
            psnk = True
        Else
            psnk = False
        End If
    End If
    ds2.Close
    pqty = InputBox("How many units are expected?", ds!pallet & " Units/pallet produced..", "10000")
    If pqty < 1 Then
        ds.Close
        Exit Sub
    End If
    
    zid = wd_seq("ProdRcv", Form1.bbsr)
    s = "INSERT INTO ProdRcv (ID, SKU, ProdDate, Units, SP_Flag, Lot_Num,"
    s = s & " RecDate1, RecDate2, RecDate3, SR1, SR2, SR3, SR4, SR5) VALUES ("      'jv042114
    s = s & zid & ","
    s = s & "'" & psku & "',"
    s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
    s = s & Val(pqty) & ","
    s = s & "'" & Val(psnk) & "',"
    plot = Right$(Format$(pdate, "m-d-yyyy"), 2)
    plot = plot & Format(DateDiff("d", DateValue("1/1/" & plot), DateValue(pdate)), "000")
    plot = Val(plot) + 1
    s = s & "'" & Format(Val(plot), "00000") & "',"
    s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
    If pdays > 1 Then
        If WeekDay(pdate) = 6 Then
            If MsgBox("Production this Saturday?", vbYesNo, "Friday Production") = vbYes Then
                s = s & "'" & Format(DateAdd("d", 1, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            Else
                s = s & "'" & Format(DateAdd("d", 3, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 4, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            End If
        Else
            If WeekDay(pdate) = 7 Then
                s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 3, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            Else
                s = s & "'" & Format(DateAdd("d", 1, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    If WeekDay(pdate) = 5 Then
                        If MsgBox("Production this Saturday?", vbYesNo, "Thursday Product") = vbYes Then
                            s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                        Else
                            s = s & "'" & Format(DateAdd("d", 4, pdate), "mm-dd-yyyy") & "',"
                        End If
                    Else
                        s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                    End If
                Else
                    s = s & "NULL,"
                End If
            End If
        End If
    Else
        s = s & "NULL,NULL,"
    End If
    i = Int(pqty / ds!pallet)
    If ds!whs_num = 1 Then
        s = s & i & ",0,0,0,0)"                                 'jv042114
    Else
        If ds!whs_num = 2 Then
            s = s & "0," & i & ",0,0,0)"                        'jv042114
        Else
            If ds!whs_num = 3 Then
                s = s & "0,0," & i & ",0,0)"                    'jv042114
            Else
                If ds!whs_num = 4 Then
                    s = s & "0,0,0," & i & ",0)"                'jv042114
                Else
                    If ds!whs_num = 5 Then                      'jv042114
                        s = s & "0,0,0,0," & i & ")"            'jv042114
                    Else
                        s = s & "0,0,0,0,0)"                    'jv042114
                    End If
                End If
            End If
        End If
    End If
    Wdb.Execute s                       'jv060916
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "instldate_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " instldate_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub opcount_Click()
    Dim ds As adodb.Recordset, s As String
    Dim rt As String, rh As String, rf As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "select opseq,oplist.sku,fgunit,fgdesc from oplist,skumast"
    s = s & " where skumast.sku = oplist.sku order by opseq,oplist.sku"
    Set ds = Sdb.Execute(s)             'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!opseq & Chr(9) & Chr(9)
            s = s & ds(1) & Chr(9)
            s = s & ds!fgunit & " " & ds!fgdesc & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Seq #||^SKU|<Description|^Pallets|^Wraps|Units"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 4000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Screen.MousePointer = 0
    
    rt = "Order Pick Count Sheet"
    rh = "Order Pick Listing"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "opcount_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " opcount_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub outstk_Click()
    Dim x As Double
    Dim ss As adodb.Recordset, s As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    s = "select sku,fgunit,fgdesc from skumast"
    s = s & " where sku in (select sku from plantskus where plant = 50"
    s = s & " and outflag = 'Y' and outstk  > 0) order by sku"
    Set ss = Sdb.Execute(s)         'jv060916
    Open Form1.tempdir & "\stock.txt" For Output As #1
    Print #1, "   Out of Stock Report     "; Format(Now, "m-d-yyyy  hh:mm Am/Pm")
    Print #1, " "
    Print #1, " "
    Print #1, "+-------------------------+"
    Print #1, "|       Brenham           |"
    Print #1, "+-------------------------+"
    Print #1, " "
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            Print #1, ss!sku & " " & ss!fgunit & " " & ss!fgdesc
            ss.MoveNext
        Loop
    End If
    Print #1, " "
    Print #1, "+-------------------------+"
    Print #1, "|      Broken Arrow       |"
    Print #1, "+-------------------------+"
    Print #1, " "
    ss.Close
    s = "select sku,fgunit,fgdesc from skumast"
    s = s & " where sku in (select sku from plantskus where plant = 51"
    s = s & " and outflag = 'Y' and outstk  > 0) order by sku"
    Set ss = Sdb.Execute(s)         'jv060916
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            Print #1, ss!sku & " " & ss!fgunit & " " & ss!fgdesc
            ss.MoveNext
        Loop
    End If
    Print #1, " "
    Print #1, "+-------------------------+"
    Print #1, "|        Sylacauga        |"
    Print #1, "+-------------------------+"
    Print #1, " "
    ss.Close
    s = "select sku,fgunit,fgdesc from skumast"
    s = s & " where sku in (select sku from plantskus where plant = 52"
    s = s & " and outflag = 'Y' and outstk  > 0) order by sku"
    Set ss = Sdb.Execute(s)         'jv060916
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            Print #1, ss!sku & " " & ss!fgunit & " " & ss!fgdesc
            ss.MoveNext
        Loop
    End If
    Close #1
    ss.Close
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\stock.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "outstk_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " outstk_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub partpalp_Click()
    partlabs.Show
End Sub

Private Sub pjobtrl_Click()
    jobtotrl.jbob = "job"
    jobtotrl.Show
End Sub

Private Sub plantotals_Click()
    plantots.Show
End Sub

Private Sub pparts_Click()
    Dim s As String, i As Integer, mdate As String, mp As String
    Dim ds As adodb.Recordset, br As adodb.Recordset, mplant As String
    Dim mbr As String
    On Error GoTo vberror
    mdate = InputBox$("Shipping Date", "Shipping Date", Form1.cdate)
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid date format....", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Form1.cdate = Format(mdate, "m-d-yyyy")
    mp = InputBox$("Plant Code", "Plant Code", Form1.plantno.Text)
    If Len(mp) = 0 Then Exit Sub
    Set ds = Sdb.Execute("select * from plants where plant = " & mp)        'jv060916
    If ds.BOF = True Then
        MsgBox "Invalid Plant Code....", vbOKOnly + vbInformation, "Sorry"
        ds.Close
        Exit Sub
    Else
        mplant = ds!plantname
    End If
    ds.Close
    mbr = "all"
    If MsgBox("Print All Branches?", vbYesNo + vbQuestion, "Select Branch") = vbNo Then
        mbr = InputBox("Branch Code:", "Enter Branch", "03")
        If Len(mbr) = 0 Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    If mbr = "all" Then
        s = "select branch,branchname from branches where branch in "
        s = s & "(select branch from brorders where orddate = '" & mdate & "'"
        s = s & " and partqty > 0 and plant = " & mp & ")"
    Else
        s = "select branch,branchname from branches where branch = " & Val(mbr)
    End If
    Set br = Sdb.Execute(s)             'jv060916
    If br.BOF = False Then
        br.MoveFirst
        Do Until br.EOF
            Printer.FontName = "MS Sans Serif"
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print " ": Printer.Print " "
            Printer.Print " ": Printer.Print " "
            Printer.Print "Partial Pallet Order"
            Printer.FontBold = False
            Printer.FontSize = 10
            Printer.Print Format(mdate, "mmmm d, yyyy"); "   "; mplant
            Printer.Print " "
            Printer.FontBold = True
            Printer.FontSize = 12
            Printer.Print br!branch & " "; br!branchname
            Printer.FontBold = False
            Printer.FontSize = 10
            Printer.Print " "
            s = "select brorders.sku,partqty,fgunit & ' ' & fgdesc from brorders,skumast"
            s = s & " Where branch = " & br!branch
            s = s & " and plant = " & mp
            s = s & " and orddate = '" & mdate & "'"
            s = s & " and partqty > 0"
            s = s & " and brorders.sku = skumast.sku"
            Set ds = Sdb.Execute(s)     'jv060916
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    Printer.Print " "
                    Printer.Print ds(0); " "; ds(2); Tab(60); ds(1)
                    ds.MoveNext
                Loop
            End If
            ds.Close
            br.MoveNext
            If br.EOF Then
                Printer.EndDoc
            Else
                Printer.NewPage
            End If
        Loop
    End If
    br.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "pparts_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " pparts_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub pprodrct_Click()
    Prodrcpts.Show
End Sub

Private Sub pro11iadj_Click()
    proadj.Show
End Sub

Private Sub procjob_Click()
    Joborders.Show
End Sub

Private Sub prodtots_Click()
    Dim pdate As String, wc As Integer, ic As Long
    Dim ds As adodb.Recordset, s As String
    Dim ssku As String, su As Long, tu As Long
    Dim rt As String, rh As String, rf As String
    On Error GoTo vberror
    pdate = InputBox("Production Date:", "Production Date")
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try again..."
        Exit Sub
    End If
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "select curr_rcpt.sku,lot_num,whse_num,pallets,wraps,description,uom_type,uom_per_pallet,qty_per_pallet"
    s = s & " from curr_rcpt,sku_config"
    s = s & " where rcpt_date = '" & pdate & "'"
    s = s & " and sku_config.sku = curr_rcpt.sku"
    s = s & " order by curr_rcpt.sku,lot_num,whse_num"
    Set ds = Wdb.Execute(s)             'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        ssku = ds(0)
        su = 0: tu = 0
        Do Until ds.EOF
            If ds(0) <> ssku Then
                s = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & su
                Grid1.AddItem s
                su = 0
                ssku = ds(0)
            End If
            s = ds(0) & Chr(9)
            s = s & ds!uom_type & " " & ds!description & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!whse_num & Chr(9)
            s = s & Format(ds!pallets, "#") & Chr(9)
            s = s & Format(ds!wraps, "#") & Chr(9)
            If ds!uom_per_pallet > 0 And ds!qty_per_pallet > 0 Then
                wc = ds!uom_per_pallet / ds!qty_per_pallet
            Else
                wc = 0
            End If
            uc = ds!pallets * ds!uom_per_pallet
            uc = uc + (ds!wraps * wc)
            s = s & uc
            su = su + uc
            tu = tu + uc
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    s = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & su
    Grid1.AddItem s
    s = Chr(9) & "Total Units" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & tu
    Grid1.AddItem s

    ds.Close
    Grid1.FormatString = "^SKU|<Description|^Lot #|^Whs|^Pallets|^Wraps|^Units"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
            
    rt = "Production Totals"
    rh = "Production Date - " & pdate
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "prodtots_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " prodtots_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub prtbbol_Click()
    blnkbill.Show
End Sub

Private Sub renamtrl_Click()
    Dim odate As String, obr As String, otrl As String
    Dim ndate As String, nbr As String, ntrl As String
    Dim ds As adodb.Recordset, sqlx As String, bagc As String
    On Error GoTo vberror
    odate = InputBox("Original Date:", "Original Date", Form1.cdate)
    If Len(odate) = 0 Then Exit Sub
    If IsDate(odate) = False Then
        MsgBox "Invalid Date format...", vbOKOnly + vbExclamation, "Aborting.."
        Exit Sub
    End If
    obr = InputBox("Original Branch Code:", "Original Branch", "3")
    If Len(obr) = 0 Or Val(obr) = 0 Then Exit Sub
    otrl = InputBox("Original Trailer Code:", "Original Trailer", "#1")
    If Len(otrl) = 0 Then Exit Sub
    sqlx = "select * from trailers where shipdate = '" & odate & "'"
    sqlx = sqlx & " and branch = " & obr
    sqlx = sqlx & " and trlno = '" & otrl & "'"
    Set ds = Sdb.Execute(sqlx)              'jv060916
    If ds.BOF = True Then
        MsgBox "Branch: " & obr & " Trailer: " & otrl & " Not found on " & odate & ".", vbOKOnly + vbExclamation, "Aborting"
        ds.Close
        Exit Sub
    End If
    ndate = InputBox("New Date:", "New Date", odate)
    If Len(ndate) = 0 Then
        ds.Close
        Exit Sub
    End If
    If IsDate(ndate) = False Then
        MsgBox "Invalid Date format...", vbOKOnly + vbExclamation, "Aborting.."
        ds.Close
        Exit Sub
    End If
    Form1.cdate = Format(ndate, "m-d-yyyy")
    nbr = InputBox("New Branch Code:", "New Branch", obr)
    If Len(nbr) = 0 Or Val(nbr) = 0 Then
        ds.Close
        Exit Sub
    End If
    ntrl = InputBox("New Trailer Code:", "New Trailer", otrl)
    If Len(ntrl) = 0 Then
        ds.Close
        Exit Sub
    End If
    ntrl = Left(ntrl, 2)
    sqlx = "OK to Rename " & odate & " Branch " & obr & " Trailer " & otrl
    sqlx = sqlx & " To: " & ndate & " Branch " & nbr & " Trailer " & ntrl
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
        ds.Close
        Exit Sub
    End If
    Screen.MousePointer = 11
    ds.MoveFirst
    Do Until ds.EOF
        sqlx = "Update trailers set shipdate = '" & ndate & "'"
        sqlx = sqlx & ", branch = " & Val(nbr)
        sqlx = sqlx & ", trlno = '" & ntrl & "'"
        If Form1.plantno <> "50" Then
            bagc = "T" & mid(Format(ndate, "mmddyyyy"), 3, 2)
            bagc = bagc & Format(Val(nbr), "00")
            bagc = bagc & Right(ntrl, 1)
            sqlx = sqlx & ", groupcode = '" & bagc & "'"
        End If
        sqlx = sqlx & " Where id = " & ds!id
        Sdb.Execute sqlx            'jv060916
        ds.MoveNext
    Loop
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "renamtrl_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " renamtrl_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub sealtrax_Click()
    sealtrak.Show
End Sub

Private Sub singmat_Click()
    Dim sqlx As String, ds As adodb.Recordset, bno As adodb.Recordset, i As Integer
    Dim x, pdate As String, pplant As String
    On Error GoTo vberror
    pdate = InputBox("Please enter order date.", "Order Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date..", vbOKOnly + vbExclamation, "Aborting Request"
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    pplant = InputBox("Please enter Plant Code.", "Plant Code", "50")
    If Len(pplant) = 0 Then Exit Sub
    Screen.MousePointer = 11
    sqlx = "Select * From Branches Where Branch in"
    sqlx = sqlx & " (Select branch From plantbranch where plant = " & pplant & ")"
    sqlx = sqlx & " order by branch"
    Set bno = Sdb.Execute(sqlx)             'jv060916
    Open Form1.tempdir & "\singpal.txt" For Output As #1
    Print #1, "Single Pallet Matches: " & Format(pdate, "m-d-yyyy")
    bno.MoveFirst
    Do Until bno.EOF
        Print #1, " "
        Print #1, "** " & bno!branchname & " **"
        Print #1, " "
        sqlx = "Select brorders.branch,branchname,count(*) from brorders,branches"
        sqlx = sqlx & " Where orddate = '" & pdate & "'"
        sqlx = sqlx & " and plant = " & pplant
        sqlx = sqlx & " And netqty in (1,3,5)"  'Mod 2 = 1"
        sqlx = sqlx & " And SKU in (Select SKU From brorders Where Branch = " & bno!branch
        sqlx = sqlx & " and orddate = '" & pdate & "'"
        sqlx = sqlx & " and plant = " & pplant
        sqlx = sqlx & " and netqty in (1,3,5))" 'Mod 2 = 1)"
        sqlx = sqlx & " And SKU in (Select SKU From Whstotals Where Whs_num in (1,2,3))"
        sqlx = sqlx & " And brorders.branch = branches.branch"
        sqlx = sqlx & " Group by brorders.branch, branchname"
        Set ds = Sdb.Execute(sqlx)          'jv060916
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds(0) <> bno("branch") Then
                    sqlx = Format(ds(0), "00") & Space(2)
                    sqlx = sqlx & ds(1)
                    sqlx = sqlx & Space(40 - Len(sqlx))
                    sqlx = sqlx & Format(ds(2), "##0")
                    Print #1, sqlx
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        bno.MoveNext
    Loop
    Close #1
    bno.Close
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\singpal.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "singmat_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " singmat_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub skucomp_Click()
    skumast1.Show
End Sub

Private Sub skuissprt_Click()
    Dim ds As adodb.Recordset, sqlx As String, pdate As String, psku As String
    Dim pdesc As String
    On Error GoTo vberror
    pdate = InputBox("Please enter issue date.", "Issue Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date..", vbOKOnly + vbExclamation, "Aborting Request"
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    psku = InputBox("SKU #:", "SKU Number...", "441")
    If Len(psku) = 0 Then Exit Sub
    Set ds = Sdb.Execute("select * from skumast where sku = '" & psku & "'")        'jv060916
    If ds.BOF Then
        MsgBox "Invalid SKU..", vbOKOnly + vbExclamation, "Aborting Request.."
        ds.Close
        Exit Sub
    End If
    pdesc = ds!fgunit & " " & ds!fgdesc
    ds.Close
    Screen.MousePointer = 11
    Open Form1.tempdir & "\skuissue.txt" For Output As #1
    Print #1, "Branch Issues - " & Format(pdate, "m-d-yyyy")
    Print #1, " "
    Print #1, "SKU: " & psku & "   " & pdesc
    Print #1, Space(33) & "  Group       Pallets  Wraps    Units"
    sqlx = "Select trailers.branch,branchname,trlno,groupcode,pallets,wraps,units"
    sqlx = sqlx & " From trailers,branches"
    sqlx = sqlx & " Where units > 0 and shipdate = '" & pdate & "'"
    sqlx = sqlx & " and sku = '" & psku & "'"
    sqlx = sqlx & " And trailers.branch = branches.branch"
    Set ds = Sdb.Execute(sqlx)          'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = Format(ds!branch, "00") & Space(1)
            sqlx = sqlx & ds!branchname & Space(1)
            sqlx = sqlx & ds!trlno
            sqlx = sqlx & Space(35 - Len(sqlx))
            sqlx = sqlx & ds!groupcode
            sqlx = sqlx & Space(50 - Len(sqlx))
            sqlx = sqlx & Format(ds!pallets, "0") '& "     "
            sqlx = sqlx & Space(58 - Len(sqlx))
            sqlx = sqlx & Format(ds!wraps, "0") '& "     "
            sqlx = sqlx & Space(66 - Len(sqlx))
            sqlx = sqlx & Format(ds!units, "0")
            Print #1, sqlx
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\skuissue.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, skuissprt.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " skuissprt_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub skuordprt_Click()
    Dim ds As adodb.Recordset, sqlx As String, pdate As String, psku As String
    Dim pdesc As String, ppl As String
    On Error GoTo vberror
    pdate = InputBox("Please enter order date.", "Order Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date..", vbOKOnly + vbExclamation, "Aborting Request"
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    psku = InputBox("SKU #:", "SKU Number...", "441")
    If Len(psku) = 0 Then Exit Sub
    Set ds = Sdb.Execute("select * from skumast where sku = '" & psku & "'")    'jv060916
    If ds.BOF Then
        MsgBox "Invalid SKU..", vbOKOnly + vbExclamation, "Aborting Request.."
        ds.Close
        Exit Sub
    End If
    pdesc = ds!fgunit & " " & ds!fgdesc
    ds.Close
    Screen.MousePointer = 11
    Open Form1.tempdir & "\skuorder.txt" For Output As #1
    Print #1, "Branch Orders - " & Format(pdate, "m-d-yyyy")
    Print #1, "SKU: " & psku & "   " & pdesc
    Print #1, Space(33) & "Order Grpd  Net   Wraps"
    sqlx = "Select brorders.branch,branchname,ordqty,grpqty,netqty,partqty,plant"
    sqlx = sqlx & " From brorders,branches"
    sqlx = sqlx & " Where ordqty > 0 and orddate = '" & pdate & "'"
    sqlx = sqlx & " and sku = '" & psku & "'"
    sqlx = sqlx & " And brorders.branch = branches.branch"
    sqlx = sqlx & " order by plant,brorders.branch"
    Set ds = Sdb.Execute(sqlx)          'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        ppl = "0"
        Do Until ds.EOF
            If ds!plant <> ppl Then
                Print #1, " "
                If ds!plant = "50" Then Print #1, "Brenham Plant - 50"
                If ds!plant = "51" Then Print #1, "Broken Arrow - 51"
                If ds!plant = "52" Then Print #1, "Sylacauga - 52"
                ppl = ds!plant
            End If
            sqlx = Format(ds!branch, "00") & Space(1)
            sqlx = sqlx & ds!branchname
            sqlx = sqlx & Space(35 - Len(sqlx))
            sqlx = sqlx & Format(ds!ordqty, "0") & "     "
            sqlx = sqlx & Format(ds!grpqty, "0") & "     "
            sqlx = sqlx & Format(ds!netqty, "0") & "     "
            sqlx = sqlx & Format(ds!partqty, "0")
            Print #1, sqlx
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\skuorder.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, skuordprt.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " skuordprt_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub nt_availsr5()
    Dim ds As adodb.Recordset, sqlx As String, gq As Integer
    Dim ds2 As adodb.Recordset, s As String, pkey As Long
    On Error GoTo vberror
    Screen.MousePointer = 11
    sqlx = "delete from whstotals where whs_num in (1,2,3,5)"
    Sdb.Execute sqlx        'jv060916
    sqlx = "select whse_num,sku,sum(qty) from lane"
    If MsgBox("Import Blocked Bays?", vbYesNo + vbQuestion, "Blocked Bays....") = vbNo Then
        sqlx = sqlx & " where lane_status <> 'B'"
        If MsgBox("Import On Hold Products?", vbYesNo + vbQuestion + vbDefaultButton2, "On Hold Product...") = vbNo Then
            sqlx = sqlx & " and lane_status <> 'H'"
        End If
    Else
        If MsgBox("Import On Hold Products?", vbYesNo + vbQuestion + vbDefaultButton2, "On Hold Product...") = vbNo Then
            sqlx = sqlx & " where lane_status <> 'H'"
        End If
    End If
    sqlx = sqlx & " group by whse_num,sku"
    sqlx = sqlx & " having sum(qty) > 0"
    Set ds2 = Wdb.Execute(sqlx)         'jv060916
    If ds2.BOF = False Then
        ds2.MoveFirst
        Do Until ds2.EOF
            pkey = wd_seq("whstotals", Form1.shipdb)
            sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail, old_qty, old_lot)"
            sqlx = sqlx & " Values (" & pkey
            sqlx = sqlx & ", " & ds2!whse_num
            sqlx = sqlx & ", '" & ds2!sku & "'"
            If ds2!whse_num = 1 Then
                sqlx = sqlx & ", " & ds2(2) - 2
            Else
                sqlx = sqlx & ", " & ds2(2)
            End If
            sqlx = sqlx & ", 0"
            If ds2!whse_num = 1 Then
                sqlx = sqlx & ", " & ds2(2) - 2
            Else
                sqlx = sqlx & ", " & ds2(2)
            End If
            sqlx = sqlx & ", 0, ' ')"
            Sdb.Execute sqlx            'jv060916
            ds2.MoveNext
        Loop
    End If
    ds2.Close
    
    sqlx = "select whse_num,sku,lot_num,lot_date,sum(qty) from lane"
    sqlx = sqlx & " where lot_date < '" & Format(DateAdd("d", -30, Now), "m-d-yyyy") & "'"
    sqlx = sqlx & " group by whse_num,sku,lot_num,lot_date"
    sqlx = sqlx & " having sum(qty) > 0"
    sqlx = sqlx & " order by whse_num,sku,lot_date"
    Set ds2 = Wdb.Execute(sqlx)                 'jv060916
    If ds2.BOF = False Then
        ds2.MoveFirst
        Do Until ds2.EOF
            sqlx = "select * from whstotals where whs_num = " & ds2!whse_num & " and sku = '" & ds2!sku & "'"
            Set ds = Sdb.Execute(sqlx)          'jv060916
            If ds.BOF = False Then
                ds.MoveFirst
                sqlx = "Update whstotals set old_qty = old_qty + " & ds2(4)
                sqlx = sqlx & ", old_lot = '" & ds2!lot_num & "' Where id = " & ds!id
                Sdb.Execute sqlx                'jv060916
            End If
            ds.Close
            ds2.MoveNext
        Loop
    End If
    
    sqlx = "select to_whse_num,sku,order_qty,ship_plt_qty from ship_infc"
    sqlx = sqlx & " where ship_status <> 'CANC' and ship_status <> 'DONE'"
    sqlx = sqlx & " and order_qty > ship_plt_qty"
    Set ds2 = Wdb.Execute(sqlx)             'jv060916
    If ds2.BOF = False Then
        ds2.MoveFirst
        Do Until ds2.EOF
            sqlx = "select * from whstotals where whs_num = " & ds2!to_whse_num & " and sku = '" & ds2!sku & "'"
            Set ds = Sdb.Execute(sqlx)      'jv060916
            If ds.BOF = False Then
                ds.MoveFirst
                gq = ds!grp_qty + (ds2!order_qty - ds2!ship_plt_qty)
                sqlx = "Update whstotals set grp_qty = " & gq
                sqlx = sqlx & ", avail = " & ds!count_qty - gq & " Where id = " & ds!id
                Sdb.Execute sqlx            'jv060916
            End If
            ds.Close
            ds2.MoveNext
        Loop
    End If
    'SR5
    sqlx = "select product,count(*) from paltasks where area = 'DOCK'"
    sqlx = sqlx & " and status = 'PEND' and lotnum < '0' and source = 'SR5'"
    sqlx = sqlx & " group by product"
    Set ds2 = Wdb.Execute(sqlx)             'jv060916
    If ds2.BOF = False Then
        ds2.MoveFirst
        Do Until ds2.EOF
            sqlx = "select * from whstotals where whs_num = 5 and sku = '" & Trim(Left(ds2!product, 4)) & "'"
            Set ds = Sdb.Execute(sqlx)      'jv060916
            If ds.BOF = False Then
                ds.MoveFirst
                gq = ds!grp_qty + ds2(1)
                sqlx = "Update whstotals set grp_qty = " & gq
                sqlx = sqlx & ", avail = " & ds!count_qty - gq & " Where id = " & ds!id
                Sdb.Execute sqlx            'jv060916
            End If
            ds.Close
            ds2.MoveNext
        Loop
    End If
    
    
    ds2.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "nt_availsr5", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " nt_availsr5 - Error Number: " & eno
        End
    End If
End Sub

Private Sub nt_ship()
    Dim ds As adodb.Recordset, ds2 As adodb.Recordset
    Dim sqlx As String, pgrp As String, pvert As Integer, phorz As Integer
    Dim pside As String, opgrp As Boolean, pkey As Long
    On Error GoTo vberror
    pgrp = UCase(InputBox$("Input Shipping Group to be posted...", "Shipping Group", Form1.cgrp))
    If Len(pgrp) = 0 Then Exit Sub
    opgrp = False
    If InStr(1, pgrp, "OP") > 0 Then
        If MsgBox("Do you wish to assign Order Pick Lanes?", vbYesNo + vbQuestion, "OP Group...") = vbYes Then opgrp = True
    End If
    Screen.MousePointer = 11
    sqlx = "Update ship_infc set ship_status = 'CANC' where order_num = '" & pgrp & "'"
    Wdb.Execute sqlx                'jv060916
    sqlx = "Update drop_infc set drop_qty = 0 where group_num = '" & pgrp & "'"
    Wdb.Execute sqlx                'jv060916
    sqlx = "select sku,shipdate,whs_num,sum(pallets) from trailers where groupcode = '" & pgrp & "'"
    sqlx = sqlx & " And pallets > 0 And sku <> 'PAR' And ra_flag = 'N' And pb_flag = 'N'"
    sqlx = sqlx & " And whs_num in (Select whs_num From Warehouses"
    sqlx = sqlx & " Where whs in ('SR1','SR2','SR3'))"
    sqlx = sqlx & " Group by sku,shipdate,whs_num"
    Set ds = Sdb.Execute(sqlx)              'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pvert = 0
            If opgrp = True Then
                sqlx = "select * from opbays where whse_num = " & ds!whs_num
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Wdb.Execute(sqlx)             'jv060916
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    pvert = ds2!vert_loc
                    phorz = ds2!horz_loc
                    pside = ds2!rack_side
                End If
                ds2.Close
            End If
            If pvert = 0 Then
                sqlx = "select * from sr_config where whs_num = " & ds!whs_num
                Set ds2 = Wdb.Execute(sqlx)             'jv060916
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    pvert = ds2!ship1_lane_vert
                    phorz = ds2!ship1_lane_horz
                    pside = ds2!ship1_lane_side
                End If
                ds2.Close
            End If
            sqlx = "select * from ship_infc where ship_status = 'CANC'"
            sqlx = sqlx & " or ship_status = 'DONE' order by id"
            Set ds2 = Wdb.Execute(sqlx)                 'jv060916
            If ds2.BOF = False Then
                sqlx = "Update ship_infc set order_num = '" & pgrp & "'"
                sqlx = sqlx & ", sku = '" & ds!sku & "'"
                sqlx = sqlx & ", lot_num = ' '"
                sqlx = sqlx & ", ship_date = '" & ds!shipdate & "'"
                sqlx = sqlx & ", order_qty = " & ds(3)
                sqlx = sqlx & ", ship_uom_qty = 0"
                sqlx = sqlx & ", ship_plt_qty = 0"
                sqlx = sqlx & ", ship_status = 'NEW'"
                sqlx = sqlx & ", to_whse_num = " & ds!whs_num
                sqlx = sqlx & ", to_vert_loc = " & pvert
                sqlx = sqlx & ", to_horz_loc = " & phorz
                sqlx = sqlx & ", to_rack_side = '" & pside & "'"
                sqlx = sqlx & ", resv_strategy = 'A'"
                sqlx = sqlx & " Where id = " & ds2!id
                Wdb.Execute sqlx            'jv060916
            Else
                pkey = wd_seq("ship_infc", Form1.bbsr)
                sqlx = "Insert into ship_infc (id, order_num, sku, lot_num, ship_date, order_qty, ship_uom_qty"
                sqlx = sqlx & ", ship_plt_qty, ship_status, to_whse_num, to_vert_loc, to_horz_loc, to_rack_side"
                sqlx = sqlx & ", resv_strategy) Values (" & pkey
                sqlx = sqlx & ", '" & pgrp & "'"
                sqlx = sqlx & ", '" & ds!sku & "'"
                sqlx = sqlx & ", ' '"
                sqlx = sqlx & ", '" & ds!shipdate & "'"
                sqlx = sqlx & ", " & ds(3)
                sqlx = sqlx & ", 0, 0, 'NEW'"
                sqlx = sqlx & ", " & ds!whs_num
                sqlx = sqlx & ", " & pvert
                sqlx = sqlx & ", " & phorz
                sqlx = sqlx & ", '" & pside & "'"
                sqlx = sqlx & ", 'A')"
                Wdb.Execute sqlx            'jv060916
            End If
            ds2.Close
            ds.MoveNext
        Loop
    Else
        MsgBox "There were no SR products found to post for this group..", vbOKOnly, "Group " & pgrp
    End If
    ds.Close
    Form1.cgrp = pgrp
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "nt_ship", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " nt_ship - Error Number: " & eno
        End
    End If
End Sub
Private Sub sndship_Click()
    Call nt_ship
End Sub

Private Sub sporders_Click()
    Dim ds As adodb.Recordset, sqlx As String, sp As Integer, pdate As String
    On Error GoTo vberror
    pdate = InputBox("Please enter order date.", "Order Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date..", vbOKOnly + vbExclamation, "Aborting Request"
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Set ds = Sdb.Execute("select * from warehouses where whs = 'SP'")  'jv060916
    If ds.BOF Then
        MsgBox "Cannot find SP warehouse in listing", vbOKOnly
        ds.Close
        Exit Sub
    End If
    sp = ds!whs_num: ds.Close
    Screen.MousePointer = 11
    Open Form1.tempdir & "\sporder.txt" For Output As #1
    Print #1, "Snack Plant Pallet Orders - " & Format(pdate, "m-d-yyyy")
    Print #1, Space(66) & "Order Grpd Net"
    sqlx = "Select brorders.branch,branchname,brorders.sku,fgunit,fgdesc,ordqty,grpqty,netqty"
    sqlx = sqlx & " From brorders,branches,skumast"
    sqlx = sqlx & " Where ordqty > 0 and orddate = '" & pdate & "'"
    sqlx = sqlx & " and brorders.plant = 50"
    sqlx = sqlx & " And brorders.Sku in (Select Sku from skumast where Whs_num = " & sp & ")"     'jv062011
    sqlx = sqlx & " And brorders.branch = branches.branch"
    sqlx = sqlx & " And brorders.sku = skumast.sku"
    Set ds = Sdb.Execute(sqlx)          'jv060916
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = Format(ds(0), "00") & Space(1)
            sqlx = sqlx & ds(1)
            sqlx = sqlx & Space(25 - Len(sqlx))
            sqlx = sqlx & ds(2) & Space(1)
            sqlx = sqlx & ds!fgunit & " " & ds!fgdesc 'ds(3)
            sqlx = sqlx & Space(70 - Len(sqlx))
            sqlx = sqlx & Format(ds!ordqty, "0") & "   "
            sqlx = sqlx & Format(ds!grpqty, "0") & "   "
            sqlx = sqlx & Format(ds!netqty, "0")
            Print #1, sqlx
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\sporder.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sporders_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sporders_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub tcycle_Click()
    Trailcycle.Show
End Sub

Private Sub trailbills_Click()
    trailbill.Show
End Sub

Private Sub trlra_Click()
    Rastrail.Show
End Sub

Private Sub trnsched_Click()
    Transched.Show
End Sub

Private Sub trnschnotes_Click()
    trucknotes.Show
End Sub

Private Sub vallists_Click()
    wdvalists.Show
End Sub

Private Sub vuettot_Click()
    Dim db As DAO.Database, ds As DAO.Recordset, sqlx As String, x
    'On Error GoTo vberror
    pdate = InputBox("Ship Date:", "Trailer Ship Date", Form1.cdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try again..."
        Exit Sub
    End If
    Screen.MousePointer = 11
    Form1.cdate = Format(pdate, "m-d-yyyy")
    Open Form1.tempdir & "\ratrtot.txt" For Output As #1
    Print #1, "             Trailer Totals - " & Format(pdate, "m-d-yyyy")
    'Print #1, " "
    If MsgBox("Brenham", vbYesNo + vbQuestion, "Brenham") = vbYes Then
        Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.shipdb)
        sqlx = "select plant,branch,account,trlno,shipdate,sum(pallets),sum(units)"
        sqlx = sqlx & " from trailers"
        sqlx = sqlx & " where plant = 50 and shipdate <= #" & pdate & "#"
        sqlx = sqlx & " and pb_flag = 'Y'"
        sqlx = sqlx & " group by plant,branch,account,trlno,shipdate"
        'MsgBox sqlx, vbOKOnly, Form1.shipdb
        Set ds = db.OpenRecordset(sqlx)
        If ds.BOF = False Then
            Print #1, "Brenham"
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = Format(ds!plant, "00") & " "
                sqlx = sqlx & Format(ds!branch, "00") & " "
                sqlx = sqlx & ds!account
                sqlx = sqlx & Space(13 - Len(sqlx))
                sqlx = sqlx & ds!trlno
                sqlx = sqlx & Space(16 - Len(sqlx))
                sqlx = sqlx & Format(ds!shipdate, "mm-dd-yyyy") & " "
                sqlx = sqlx & Space(8 - Len(Format(ds(5), "#####0")))
                sqlx = sqlx & Format(ds(5), "#####0")
                sqlx = sqlx & Space(8 - Len(Format(ds(6), "###,##0")))
                sqlx = sqlx & Format(ds(6), "###,##0")
                Print #1, sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    
    If MsgBox("Broken Arrow", vbYesNo + vbQuestion, "Broken Arrow") = vbYes Then
        If Right(Form1.baship, 4) = ".mdb" Then
            Set db = OpenDatabase(Form1.baship)
            sqlx = "select plant,branch,account,trlno,shipdate,sum(pallets),sum(units)"
            sqlx = sqlx & " from trailers"
            sqlx = sqlx & " where plant = 51 and shipdate <= #" & pdate & "#"
            sqlx = sqlx & " and pb_flag = true"
            sqlx = sqlx & " group by plant,branch,account,trlno,shipdate"
        Else
            Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.baship)
            sqlx = "select plant,branch,account,trlno,shipdate,sum(pallets),sum(units)"
            sqlx = sqlx & " from trailers"
            sqlx = sqlx & " where plant = 51 and shipdate <= #" & pdate & "#"
            sqlx = sqlx & " and pb_flag = 'Y'"
            sqlx = sqlx & " group by plant,branch,account,trlno,shipdate"
        End If
        Set ds = db.OpenRecordset(sqlx)
        If ds.BOF = False Then
            Print #1, "Broken Arrow"
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = Format(ds!plant, "00") & " "
                sqlx = sqlx & Format(ds!branch, "00") & " "
                sqlx = sqlx & ds!account
                sqlx = sqlx & Space(13 - Len(sqlx))
                sqlx = sqlx & ds!trlno
                sqlx = sqlx & Space(16 - Len(sqlx))
                sqlx = sqlx & Format(ds!shipdate, "mm-dd-yyyy") & " "
                sqlx = sqlx & Space(8 - Len(Format(ds(5), "#####0")))
                sqlx = sqlx & Format(ds(5), "#####0")
                sqlx = sqlx & Space(8 - Len(Format(ds(6), "###,##0")))
                sqlx = sqlx & Format(ds(6), "###,##0")
                Print #1, sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    
    If MsgBox("Sylacauga", vbYesNo + vbQuestion, "Sylacauga") = vbYes Then
        If Right(Form1.syship, 4) = ".mdb" Then
            Set db = OpenDatabase(Form1.syship)
            sqlx = "select plant,branch,account,trlno,shipdate,sum(pallets),sum(units)"
            sqlx = sqlx & " from trailers"
            sqlx = sqlx & " where plant = 52 and shipdate <= #" & pdate & "#"
            sqlx = sqlx & " and pb_flag = true"
            sqlx = sqlx & " group by plant,branch,account,trlno,shipdate"
        Else
            Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.syship)
            sqlx = "select plant,branch,account,trlno,shipdate,sum(pallets),sum(units)"
            sqlx = sqlx & " from trailers"
            sqlx = sqlx & " where plant = 52 and shipdate <= #" & pdate & "#"
            sqlx = sqlx & " and pb_flag = 'Y'"
            sqlx = sqlx & " group by plant,branch,account,trlno,shipdate"
        End If
        Set ds = db.OpenRecordset(sqlx)
        If ds.BOF = False Then
            Print #1, "Sylacauga"
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = Format(ds!plant, "00") & " "
                sqlx = sqlx & Format(ds!branch, "00") & " "
                sqlx = sqlx & ds!account
                sqlx = sqlx & Space(13 - Len(sqlx))
                sqlx = sqlx & ds!trlno
                sqlx = sqlx & Space(16 - Len(sqlx))
                sqlx = sqlx & Format(ds!shipdate, "mm-dd-yyyy") & " "
                sqlx = sqlx & Space(8 - Len(Format(ds(5), "#####0")))
                sqlx = sqlx & Format(ds(5), "#####0")
                sqlx = sqlx & Space(8 - Len(Format(ds(6), "###,##0")))
                sqlx = sqlx & Format(ds(6), "###,##0")
                Print #1, sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    Close #1
    Screen.MousePointer = 0
    x = Shell("notepad.exe " & Form1.tempdir & "\ratrtot.txt", vbNormalFocus)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "vuettot_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " vuettot_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub wdbrowstat_Click()
    browstat.Show
End Sub

Private Sub whssum_Click()
    Dim ds As adodb.Recordset, sqlx As String, x
    On Error GoTo vberror
    sqlx = "select whstotals.whs_num,warehouses.whsname,sum(avail)"
    sqlx = sqlx & " from warehouses,whstotals"
    sqlx = sqlx & " where whstotals.whs_num = warehouses.whs_num"
    sqlx = sqlx & " group by whstotals.whs_num,warehouses.whsname"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Open Form1.tempdir & "\whssum.txt" For Output As #1
        Print #1, "Warehouse Available Pallet Summary"
        Print #1, Format(Now, "m-d-yyyy h:mm Am/Pm")
        Print #1, " "
        Do Until ds.EOF
            sqlx = Format(ds(0), "00") & " "
            sqlx = sqlx & ds(1)
            sqlx = sqlx & Space(30 - Len(sqlx))
            sqlx = sqlx & Space(8 - Len(Format(ds(2), "#####0")))
            sqlx = sqlx & Format(ds(2), "#####0")
            Print #1, sqlx
            ds.MoveNext
        Loop
        Close #1
        x = Shell("notepad.exe " & Form1.tempdir & "\whssum.txt", vbNormalFocus)
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "whssum_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " whssum_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub xitmenu_Click()
    Dim i As Integer, f As String
    Dim t As Long, l As Long, h As Long, w As Long
    Wdb.Close
    Sdb.Close
    If Form1.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "form1" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Form1.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Form1.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Form1.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Form1.Width
                Exit For
            End If
        Next i
    End If
    Open localAppDataPath & "\shpforms.ini" For Output As #1
    For i = 1 To FrmGrid.Rows - 1
        f = FrmGrid.TextMatrix(i, 0)
        t = Val(FrmGrid.TextMatrix(i, 1))
        l = Val(FrmGrid.TextMatrix(i, 2))
        h = Val(FrmGrid.TextMatrix(i, 3))
        w = Val(FrmGrid.TextMatrix(i, 4))
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
