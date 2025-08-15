VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimpbanner 
   BackColor       =   &H00C0C0C0&
   Caption         =   "B.I.M.P."
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   9480
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2355
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imports"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sales Analysis"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   6600
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Planned Distribution"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   8295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "08.10.2022"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "Branch Inventory Management Program  --ds '75--"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Dale Sommerlatte Fightin' Texas Aggie Class of '75"
      Top             =   2520
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      Caption         =   "B.I.M.P."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   98.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "bimpbanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    plandist.Show
End Sub

Private Sub Command2_Click()
    whssales.Show
End Sub

Private Sub Command3_Click()
    bimptstimp.Show
End Sub

Private Sub Command4_Click()
    export_branchbarcodes_ships
    'export_branchbarcodes_bills
End Sub

Private Sub Form_Load()
    Dim s As String, ds As Recordset
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    s = "select listdisplay from valuelists where listname = 'wdbimpuser'"
    s = s & " and listreturn = '" & bimpuserid & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Command1.Visible = True
        Command2.Visible = True
        If bimpuserid = "jvierus" Or bimpuserid = "rlhalfmann" Then Command3.Visible = True
    Else
        Command2.Visible = True
    End If
    ds.Close
End Sub

Private Sub Form_Resize()
    Label1.Width = Me.Width
    Label2.Width = Me.Width
    Command1.Width = Me.Width
    Command2.Width = Me.Width
    pgrid.Width = Me.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wdb.Close
    tsb.Close
    If r12access = True Then r12db.Close
    MsgBox "Bye and happy trails!", vbOKOnly + vbExclamation, "B.I.M.P.  " & bimpuserid
    End
End Sub

