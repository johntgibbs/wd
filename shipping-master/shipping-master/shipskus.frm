VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Skulist 
   Caption         =   "Product Listing"
   ClientHeight    =   8235
   ClientLeft      =   3030
   ClientTop       =   1845
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   11805
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   5880
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
      _Version        =   327680
      ForeColor       =   16384
      BackColorFixed  =   16777152
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Information "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5655
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   19
         Left            =   1320
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   7080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
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
         Index           =   18
         Left            =   1320
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   17
         Left            =   1320
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   16
         Left            =   1320
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   15
         Left            =   1320
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   14
         Left            =   1320
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   13
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Index           =   12
         Left            =   1320
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   11
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   9
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   8
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   7
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   6
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   4800
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   4080
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "SKU"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Display Date:"
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
         Index           =   19
         Left            =   120
         TabIndex        =   40
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wrap Conv:"
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
         Index           =   18
         Left            =   120
         TabIndex        =   42
         Top             =   6720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit Lbs:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   39
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "On Hand:"
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
         Index           =   16
         Left            =   120
         TabIndex        =   38
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bundle:"
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
         Index           =   15
         Left            =   120
         TabIndex        =   37
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulk:"
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
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UPC:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   35
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pallet Conv:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G/L:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gallonage:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Invoice #:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sales Class:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Class:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Type:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RA Desc:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Warehouse:"
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
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Source:"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit:"
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
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flavor:"
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
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu inssku 
         Caption         =   "Add SKU"
      End
      Begin VB.Menu delsku 
         Caption         =   "Delete SKU"
      End
   End
End
Attribute VB_Name = "Skulist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edfield As String, edbbsr As String
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    sqlx = "select sku,fgdesc,fgunit from skumast order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!sku & Chr(9)
            sqlx = sqlx & ds!fgdesc & Chr(9)
            sqlx = sqlx & ds!fgunit
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Name|^Unit"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 3300
    Grid1.ColWidth(2) = 1500
    Screen.MousePointer = 0
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
Private Sub recupdate()
    Dim ds As adodb.Recordset
    Dim sqlx As String, i As Integer
    On Error GoTo vberror
    sqlx = "select * from skumast where sku = '" & Text1(0).Text & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If edfield = "fgdesc" Then
            Text1(1).Text = Left(Text1(1).Text, 25)
            Grid1.TextMatrix(Grid1.Row, 1) = Text1(1).Text
            sqlx = "Update skumast set fgdesc = '" & Text1(1).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "fgunit" Then
            Text1(2).Text = Left(Text1(2).Text, 12)
            Grid1.TextMatrix(Grid1.Row, 2) = Text1(2).Text
            sqlx = "Update skumast set fgunit = '" & Text1(2).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "psource" Then
            Text1(3).Text = Val(Text1(3).Text)
            sqlx = "Update skumast set psource = " & Val(Text1(3).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "whs_num" Then
            Text1(4).Text = Val(Text1(4).Text)
            sqlx = "Update skumast set whs_num = " & Val(Text1(4).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "proddesc" Then
            Text1(5).Text = Left(Text1(5).Text, 16)
            sqlx = "Update skumast set proddesc = '" & Text1(5).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "prodtype" Then
            Text1(6).Text = Left(Text1(6).Text, 1)
            sqlx = "Update skumast set prodtype = '" & Text1(6).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "prodclass" Then
            Text1(7).Text = Left(Text1(7).Text, 2)
            sqlx = "Update skumast set prodclass = '" & Text1(7).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "sales_class" Then
            Text1(8).Text = Left(Text1(8).Text, 2)
            sqlx = "Update skumast set sales_class = '" & Text1(8).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "invoice_no" Then
            Text1(9).Text = Left(Text1(9).Text, 3)
            sqlx = "Update skumast set invoice_no = '" & Text1(9).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "gal_divisor" Then
            Text1(10).Text = Val(Text1(10).Text)
            sqlx = "Update skumast set gal_divisor = " & Val(Text1(10).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "gl_number" Then
            Text1(11).Text = Val(Text1(11).Text)
            sqlx = "Update skumast set gl_number = " & Val(Text1(11).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "pallet" Then
            Text1(12).Text = Val(Text1(12).Text)
            sqlx = "Update skumast set pallet = " & Val(Text1(12).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "upc" Then
            Text1(13).Text = Left(Text1(13).Text, 12)
            sqlx = "Update skumast set upc = '" & Text1(13).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "bulku" Then
            Text1(14).Text = Val(Text1(14).Text)
            sqlx = "Update skumast set bulku = " & Val(Text1(14).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "bundle" Then
            Text1(15).Text = Val(Text1(15).Text)
            sqlx = "Update skumast set bundle = " & Val(Text1(15).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "onhand" Then
            Text1(16).Text = Val(Text1(16).Text)
            sqlx = "Update skumast set onhand = " & Val(Text1(16).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "unlbs" Then
            Text1(17).Text = Val(Text1(17).Text)
            sqlx = "Update skumast set unlbs = " & Val(Text1(17).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "numwrap" Then
            Text1(18).Text = Val(Text1(18).Text)
            sqlx = "Update skumast set numwrap = " & Val(Text1(18).Text)
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
        If edfield = "displaydate" Then
            sqlx = "Update skumast set display_date = '" & Text1(19).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Sdb.Execute sqlx
        End If
    End If
    ds.Close
    edfield = ""
    If Len(edbbsr) > 0 Then Call srupdate
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "recupdate", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " recupdate - Error Number: " & eno
        End
    End If
End Sub
Private Sub srupdate()
    Dim ds As adodb.Recordset
    Dim sqlx As String, i As Integer
    On Error GoTo vberror
    sqlx = "select * from sku_config where sku = '" & Text1(0).Text & "'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If edbbsr = "description" Then
            sqlx = "Update sku_config set description = '" & Text1(1).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Wdb.Execute sqlx
        End If
        If edbbsr = "uom_type" Then
            sqlx = "Update sku_config set uom_type = '" & Text1(2).Text & "'"
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Wdb.Execute sqlx
        End If
        If edbbsr = "uom_per_pallet" Then
            sqlx = "Update sku_config set uom_per_pallet = " & Text1(12).Text
            sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
            Wdb.Execute sqlx
            If Val(Text1(18).Text) > 0 Then
                i = Val(Text1(12).Text) / Val(Text1(18).Text)
                sqlx = "Update sku_config set qty_per_pallet = " & i
                sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
                Wdb.Execute sqlx
            End If
        End If
        If edbbsr = "qty_per_pallet" Then
            If Val(Text1(12).Text) > 0 Then
                i = Val(Text1(12).Text) / Val(Text1(18).Text)
                sqlx = "Update sku_config set qty_per_pallet = " & i
                sqlx = sqlx & " Where sku = '" & Text1(0).Text & "'"
                Wdb.Execute sqlx
            End If
        End If
    End If
    ds.Close
    edbbsr = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "srupdate", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " srupdate - Error Number: " & eno
        End
    End If
End Sub
Private Sub delrec()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "Ok to delete SKU: " & Grid1.TextMatrix(Grid1.Row, 0)
    sqlx = sqlx & " " & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
    sqlx = "select * from skumast where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Delete from skumast where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Sdb.Execute sqlx
    End If
    ds.Close
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
        Call Grid1_RowColChange
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "delrec", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " delrec - Error Number: " & eno
        End
    End If
End Sub
Private Sub insrec()
    Dim nb As String, i As Integer, nz As String, zid As Long
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    If Len(edfield) > 0 Then Call recupdate
    nb = InputBox("New SKU Code: ", "Insert new product...", "000")
    If Len(nb) = 0 Or Val(nb) = 0 Or Val(nb) > 9999 Then Exit Sub           'jv082415
    For i = 0 To Grid1.Rows - 1
        If Val(nb) = Val(Grid1.TextMatrix(i, 0)) Then
            sqlx = "SKU Number " & nb
            sqlx = sqlx & " already in use for "
            sqlx = sqlx & Grid1.TextMatrix(i, 2) & " "
            sqlx = sqlx & Grid1.TextMatrix(i, 1) & "."
            MsgBox sqlx, vbOKOnly + vbExclamation, "Sorry, cannot add..."
            Exit Sub
        End If
    Next i
    nb = Format(Val(nb), "000")
    sqlx = "Insert into skumast (sku, psource, whs_num) Values ('" & nb & "', 1, 0)"
    Sdb.Execute sqlx
    sqlx = "select * from sku_config where sku = '" & nb & "'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = True Then
        s = "INSERT INTO SKU_Config (SKU, SKU_Type, Select_Method)"
        s = s & " VALUES ('" & nb & "', 'F', 'A')"
        Wdb.Execute s
    End If
    ds.Close
    If Val(Form1.plantno) = 50 Then
        sqlx = "select * from zone_config where sku = '" & nb & "'"
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = True Then
            nz = InputBox("Zone Number For Crane:", "Zone....", "2")
            If Len(nz) = 0 Then nz = "2"
            s = "INSERT INTO Zone_Config ("
            zid = wd_seq("Zone_Config", Form1.bbsr)
            s = "INSERT INTO Zone_Config (ID, SKU, Whse_Num, Zone_Num, Lot_Size)"
            s = s & " VALUES (" & zid & ","
            s = s & "'" & nb & "',3," & Val(nz) & ",0)"
            Wdb.Execute s
        End If
        ds.Close
    End If
    
    Call refresh_grid
    For i = 0 To Grid1.Rows - 1
        If Val(nb) = Val(Grid1.TextMatrix(i, 0)) Then
            Grid1.Row = i: Grid1.TopRow = i
            Exit For
        End If
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "insrec", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " insrec - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    If Val(Left(Combo1, 2)) <> Val(Text1(3).Text) Then
        If Len(edfield) > 0 Then Call recupdate
        Text1(3).Text = Left(Combo1, 2)
        edfield = "psource"
        Call recupdate
    End If
End Sub

Private Sub Combo2_Click()
    If Val(Left(Combo2, 2)) <> Val(Text1(4).Text) Then
        If Len(edfield) > 0 Then Call recupdate
        Text1(4).Text = Left(Combo2, 2)
        edfield = "whs_num"
        Call recupdate
    End If
End Sub

Private Sub delsku_Click()
    Call delrec
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edfield) > 0 Then Call recupdate
    If skulist.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "skulist" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = skulist.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = skulist.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = skulist.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = skulist.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        KeyCode = 0: Call insrec
    End If
    If KeyCode = 120 Then
        KeyCode = 0: Call delrec
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    On Error GoTo vberror
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "skulist" Then
            Form1.FrmGrid.Col = 1: skulist.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: skulist.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: skulist.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: skulist.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Combo1.Clear: Combo2.Clear
    sqlx = "select * from prodsources order by source"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = Format(ds!source, "00") & "-" & ds!sourcename
            Combo1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.AddItem "00-Undefined"
    Combo2.AddItem "00-Any Crane"
    sqlx = "select * from warehouses where plant = " & Form1.plantno
    sqlx = sqlx & " order by whs_num"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = Format(ds!whs_num, "00") & "-" & ds!whsname
            Combo2.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
    If Grid1.Rows > 1 Then
        Grid1.Row = 1
        Call Grid1_RowColChange
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
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 680
    If Me.Width > Frame1.Width + 500 Then Grid1.Width = Me.Width - (Frame1.Width + 310)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    On Error GoTo vberror
    If Len(edfield) > 0 Then Call recupdate
    If Grid1.Row = 0 Then Exit Sub
    If Val(Text1(0).Text) <> Val(Grid1.TextMatrix(Grid1.Row, 0)) Then
        Text1(0).Text = Grid1.TextMatrix(Grid1.Row, 0)
        Text1(1).Text = "": Text1(2).Text = "": Text1(3).Text = ""
        Text1(4).Text = "": Text1(5).Text = "": Text1(6).Text = ""
        Text1(7).Text = "": Text1(8).Text = "": Text1(9).Text = ""
        Text1(10).Text = "": Text1(11).Text = "": Text1(12).Text = ""
        Text1(13).Text = "": Text1(14).Text = "": Text1(15).Text = ""
        Text1(16).Text = "": Text1(17).Text = "": Text1(18).Text = ""
        Text1(19).Text = ""
        sqlx = "select * from skumast where sku = '"
        sqlx = sqlx & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            If Len(ds!fgdesc) > 0 Then Text1(1).Text = ds!fgdesc
            If Len(ds!fgunit) > 0 Then Text1(2).Text = ds!fgunit
            If Len(ds!psource) > 0 Then Text1(3).Text = ds!psource
            If Len(ds!whs_num) > 0 Then Text1(4).Text = ds!whs_num
            If Len(ds!proddesc) > 0 Then Text1(5).Text = ds!proddesc
            If Len(ds!prodtype) > 0 Then Text1(6).Text = ds!prodtype
            If Len(ds!prodclass) > 0 Then Text1(7).Text = ds!prodclass
            If Len(ds!sales_class) > 0 Then Text1(8).Text = ds!sales_class
            If Len(ds!invoice_no) > 0 Then Text1(9).Text = ds!invoice_no
            If Len(ds!gal_divisor) > 0 Then Text1(10).Text = ds!gal_divisor
            If Len(ds!gl_number) > 0 Then Text1(11).Text = ds!gl_number
            If Len(ds!pallet) > 0 Then Text1(12).Text = ds!pallet
            If Len(ds!upc) > 0 Then Text1(13).Text = ds!upc
            If Len(ds!bulku) > 0 Then Text1(14).Text = ds!bulku
            If Len(ds!bundle) > 0 Then Text1(15).Text = ds!bundle
            If Len(ds!onhand) > 0 Then Text1(16).Text = ds!onhand
            If Len(ds!unlbs) > 0 Then Text1(17).Text = ds!unlbs
            If Len(ds!numwrap) > 0 Then Text1(18).Text = ds!numwrap
            If Len(ds!display_date) > 0 Then Text1(19).Text = ds!display_date
        End If
        ds.Close
    End If
    For i = 0 To Combo1.ListCount - 1
        If Val(Text1(3).Text) = Val(Left(Combo1.List(i), 2)) Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
    For i = 0 To Combo2.ListCount - 1
        If Val(Text1(4).Text) = Val(Left(Combo2.List(i), 2)) Then
            Combo2.ListIndex = i
            Exit For
        End If
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "grid1_rowcolchange", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " grid1_rowcolchange - Error Number: " & eno
        End
    End If
End Sub

Private Sub inssku_Click()
    Call insrec
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If Index = 0 Then edfield = "sku"
    If Index = 1 Then
        edfield = "fgdesc": edbbsr = "description"
    End If
    If Index = 2 Then
        edfield = "fgunit": edbbsr = "uom_type"
    End If
    If Index = 3 Then edfield = "psource"
    If Index = 4 Then edfield = "whs_num"
    If Index = 5 Then edfield = "proddesc"
    If Index = 6 Then edfield = "prodtype"
    If Index = 7 Then edfield = "prodclass"
    If Index = 8 Then edfield = "sales_class"
    If Index = 9 Then edfield = "invoice_no"
    If Index = 10 Then edfield = "gal_divisor"
    If Index = 11 Then edfield = "gl_number"
    If Index = 12 Then
        edfield = "pallet": edbbsr = "uom_per_pallet"
    End If
    If Index = 13 Then edfield = "upc"
    If Index = 14 Then edfield = "bulku"
    If Index = 15 Then edfield = "bundle"
    If Index = 16 Then edfield = "onhand"
    If Index = 17 Then edfield = "unlbs"
    If Index = 18 Then
        edfield = "numwrap": edbbsr = "qty_per_pallet"
    End If
    If Index = 19 Then edfield = "displaydate"
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Len(edfield) > 0 Then Call recupdate
End Sub

