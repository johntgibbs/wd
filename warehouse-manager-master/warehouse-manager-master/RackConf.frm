VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Rack Configuration"
   ClientHeight    =   11775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11805
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   11775
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   375
      Left            =   10080
      TabIndex        =   45
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Edit On-Hold Units"
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
      Left            =   2280
      TabIndex        =   44
      Top             =   9600
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1335
      Left            =   0
      TabIndex        =   43
      Top             =   10440
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2355
      _Version        =   327680
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Change BC OPCode"
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
      TabIndex        =   42
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Check Shipping"
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
      TabIndex        =   40
      Top             =   9120
      Width           =   2055
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
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text10 
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
      Left            =   9840
      TabIndex        =   35
      Text            =   "Text10"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text9 
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
      Left            =   7680
      TabIndex        =   34
      Text            =   "Text9"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "UnSplit Rack"
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
      Left            =   6840
      TabIndex        =   33
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Mark On Hold"
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
      Left            =   4920
      TabIndex        =   32
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Wrap/Un-Wrap"
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
      Left            =   2280
      TabIndex        =   31
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New Position"
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
      TabIndex        =   30
      Top             =   8640
      Width           =   2055
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
      Left            =   2280
      TabIndex        =   27
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Pallet"
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
      TabIndex        =   26
      Top             =   8160
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6735
      Left            =   0
      TabIndex        =   25
      Top             =   1320
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11880
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12648447
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.CheckBox Check2 
      Caption         =   "On Hold"
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
      Left            =   3960
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1st Out"
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
      Left            =   2880
      TabIndex        =   20
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F9: Clear Rack"
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
      Left            =   7800
      TabIndex        =   23
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F2: Accept Changes"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   840
      Width           =   1935
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
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text7 
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
      Left            =   1080
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   840
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reservation "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   10575
      Begin VB.TextBox Text6 
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
         Left            =   6000
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text5 
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
         Left            =   720
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "BB Qty:"
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
         TabIndex        =   37
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "4Way Qty:"
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
         Left            =   8760
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   " "
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
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Lot:"
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
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Inventory "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   0
      Top             =   9600
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text4 
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
         Left            =   4560
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
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
         Left            =   2520
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   720
         Width           =   615
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   720
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   720
         Width           =   855
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
         Left            =   720
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   " "
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
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "4Way Qty:"
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
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "BB Qty:"
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
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Lot:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label slit 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   41
      Top             =   9240
      Width           =   9255
   End
   Begin VB.Label Label13 
      Caption         =   "SKU Lookup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9960
      TabIndex        =   38
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label posdesc 
      Caption         =   "posdesc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6120
      TabIndex        =   29
      Top             =   8280
      Width           =   5895
   End
   Begin VB.Label possku 
      Caption         =   "possku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label RKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   8520
      TabIndex        =   24
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Slot:"
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
      Left            =   1680
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Capacity:"
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
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Menu viewmenu 
      Caption         =   "View"
      Visible         =   0   'False
      Begin VB.Menu palhist 
         Caption         =   "View Pallet History"
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function order_pick_position(osku As String) As Long
    Dim ds As ADODB.Recordset, s As String
    s = "select * from oplist where sku = '" & osku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        order_pick_position = ds!opseq
    Else
        order_pick_position = 0
    End If
    ds.Close
End Function

Private Sub ba120517()      'Change select rack capacity from 16 to 12      jv120517
    Dim ds As ADODB.Recordset, s As String
    Dim maisle As String, mrack As String                   'jv062111
    Dim rno As Long, i As Integer
    maisle = Left(RKey.Caption, 1)                          'jv062111
    mrack = Right(RKey.Caption, Len(RKey.Caption) - 2)      'jv062111
    s = "select id from racks where aisle = '" & maisle & "' and rack = '" & mrack & "'"
    MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rno = ds!id
    End If
    ds.Close
    s = "update racks set capacity = 12 where id = " & rno
    MsgBox s
    Wdb.Execute s
    s = "delete from rackpos where rackno = " & rno & " and posn_num in (1, 2, 3, 4)"
    Wdb.Execute s
    MsgBox s
    For i = 1 To 12
        s = "update rackpos set posn_num = " & i & " where rackno = " & rno & " and posn_num = " & i + 4
        MsgBox s
        Wdb.Execute s
    Next i
End Sub

Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String
    Dim maisle As String, mrack As String                   'jv062111
    Dim rno As Long, lsku As String
    Dim p As ptask
    maisle = Left(RKey.Caption, 1)                          'jv062111
    mrack = Right(RKey.Caption, Len(RKey.Caption) - 2)      'jv062111
    Text7.Enabled = True
    Text9 = " "
    Text10 = " "
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15
    s = "select * from rackpos where rackno in ("
    s = s & "select id from racks where aisle = '" & maisle & "'"
    s = s & " and rack = '" & mrack & "')"
    
    If UCase(Left(Me.Caption, 6)) = "RACK M" Then ' Or Val(Text7) = 0 Then
        s = s & " order by sku,recv_date"
        Command5.Enabled = True
        Command6.Enabled = True
        Command2.Enabled = False
    Else
        s = s & " order by recv_date,barcode"
        Command5.Enabled = False
        Command6.Enabled = False
        Command2.Enabled = True
    End If
    If maisle = "M" Then
        Label13.Visible = True
        Combo1.Visible = True
        Combo1.Clear
    Else
        Label13.Visible = False
        Combo1.Visible = False
    End If
    Command8.Visible = False: Command8.Enabled = False
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rno = ds!RackNo
        Do Until ds.EOF
            If ds!RackNo <> rno Then
                Command8.Visible = True
                Command8.Enabled = True
            End If
            s = ds!id & Chr(9)
            s = s & ds!posn_num & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!pallet_num & Chr(9)
            s = s & ds!count_qty & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!qty2 & Chr(9)
            s = s & Format(ds!recv_date, "m-d-yyyy") & Chr(9)
            'If ds!count_qty > 0 Then Text7.Enabled = False
            If ds!count_qty > 0 Then
                If ds!bbc = "N" Then
                    Text10 = Val(Text10) + 1
                    s = s & "Y"
                Else
                    Text9 = Val(Text9) + 1
                End If
            End If
            s = s & Chr(9) & ds!barcode & Chr(9)
            If ds!wrapped = "Y" Then s = s & "Y"
            s = s & Chr(9)
            'If ds!hold = "Y" Then s = s & "Y"                  'jv040815
            s = s & Chr(9) & ds!RackNo
            If UCase(Left(Me.Caption, 6)) <> "RACK M" Then
                If ds!sku > "0" Then
                    i = Val(ds!sku)
                    s = s & Chr(9) & StrConv(skurec(i).prodname, vbProperCase)
                End If
            End If
            Grid1.AddItem s
            If ds!sku > "0" And ds!sku <> lsku Then
                lsku = ds!sku
                Combo1.AddItem lsku
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 And mrack <> "OP" Then                    'jv040715
        For i = 1 To Grid1.Rows - 1                             'jv040715
            If Grid1.TextMatrix(i, 10) > "0" Then               'jv040715
                p.palletid = Grid1.TextMatrix(i, 10)            'jv040715
                p.lotnum = Grid1.TextMatrix(i, 3)               'jv040715
                p.lotnum2 = Grid1.TextMatrix(i, 6)              'jv040715
                If check_hold(p) = True Then                    'jv040715
                    Grid1.TextMatrix(i, 12) = "Y"               'jv040715
                End If                                          'jv040715
            End If                                              'jv040715
        Next i                                                  'jv040715
    End If                                                      'jv040715
        
    s = "select qty,qty4 from racks where aisle = '" & maisle & "'"
    s = s & " and rack = '" & mrack & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF Then
        ds.MoveFirst
        If ds!qty <> Val(Text9) Then
            Text9.ForeColor = posdesc.ForeColor
        Else
            Text9.ForeColor = Text3.ForeColor
        End If
        If ds!qty4 <> Val(Text10) Then
            Text10.ForeColor = posdesc.ForeColor
        Else
            Text10.ForeColor = Text3.ForeColor
        End If
    End If
    ds.Close
    possku.Caption = "..."
    posdesc.Caption = "..."
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 2) > ".." Then
                Grid1.Row = i
                Grid1_RowColChange
                Exit For
            End If
        Next i
    End If
    If UCase(Left(Me.Caption, 6)) = "RACK M" Then 'Or Val(Text7) = 0 Then
        If Grid1.Rows > 1 Then
            Grid1.FillStyle = flexFillRepeat
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 10) > "0" And Grid1.TextMatrix(i, 11) <> "Y" Then
                    Grid1.Row = i: Grid1.RowSel = i
                    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                    Grid1.CellBackColor = RKey.BackColor
                End If
            Next i
            Grid1.Row = 1
        End If
    End If
                
    If UCase(Left(Me.Caption, 6)) = "RACK M" Then 'Or Val(Text7) = 0 Then
        Grid1.FormatString = "^Id|^Position|^SKU|^Lot|^Pallet#|^Units|^Lot2|^Qty2|^Date|^4Way|^Bar Code|^Wrapped|^OnHold"
        Grid1.ColWidth(0) = 800
        Grid1.ColWidth(1) = 800
        Grid1.ColWidth(11) = 800
        Grid1.ColWidth(14) = 1
    Else
        If Command8.Visible = True Then
            Grid1.FormatString = "^Id||^SKU|^Lot|^Pallet#|^Units|^Lot2|^Qty2|^Date|^4Way|^Bar Code||^OnHold|^RackId|<Product"
        Else
            Grid1.FormatString = "^Id||^SKU|^Lot|^Pallet#|^Units|^Lot2|^Qty2|^Date|^4Way|^Bar Code||^OnHold||<Product"
        End If
        Grid1.ColWidth(0) = 800
        Grid1.ColWidth(1) = 1 '800
        Grid1.ColWidth(11) = 1
        Grid1.ColWidth(14) = 2600
    End If
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1700
    
    Grid1.ColWidth(12) = 800
    If Command8.Visible = True Then
        Grid1.ColWidth(13) = 800
    Else
        Grid1.ColWidth(13) = 1
    End If
    Text7 = Grid1.Rows - 1
    Grid1.Redraw = True
    If maisle = "M" And Combo1.ListCount > 1 Then Combo1.ListIndex = 0
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(Grid1.TextMatrix(Grid1.Row, 10), 13)
    tktonhand.bbarcode = s
    tktonhand.bproduct = posdesc
    tktonhand.Show
End Sub

Private Sub Combo1_Click()
    Dim i As Integer
    If Combo1 > "0" Then
        For i = 0 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 2) = Combo1 Then
                Grid1.TopRow = i
                Grid1.Row = i
                DoEvents
                Grid1.Col = 2
                Grid1.ColSel = Grid1.Cols - 1
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Command1_Click()
    Dim sqlx As String, y As Integer
    Dim ds As ADODB.Recordset, rid As Long
    Dim newcap As Integer, oldcap As Integer
    Dim maisle As String, mrack As String                   'jv062111
    maisle = Left(RKey.Caption, 1)                          'jv062111
    mrack = Right(RKey.Caption, Len(RKey.Caption) - 2)      'jv062111
    If Len(Text1) = 0 Then Text1 = " "
    If Len(Text2) = 0 Then Text2 = " "
    If Len(Text5) = 0 Then Text5 = " "
    If Len(Text6) = 0 Then Text6 = " "
    ' Update the header record for this rack.
    sqlx = "Update racks set qty = " & Val(Text9)
    sqlx = sqlx & ", qty4 = " & Val(Text10)
    sqlx = sqlx & ", resv_sku = '" & Text5 & "'"
    sqlx = sqlx & ", resv_lot = '" & Text6 & "'"
    sqlx = sqlx & ", slot = " & Val(Text8)
    sqlx = sqlx & ", fo = " & Check1.Value
    sqlx = sqlx & ", hold = " & Check2.Value
    sqlx = sqlx & ", Capacity = " & Val(Text7)
    sqlx = sqlx & " Where aisle = '" & maisle & "'"
    sqlx = sqlx & " and rack = '" & mrack & "'"
    Wdb.Execute sqlx
    ' Update the detail records for this rack, if the capacity has decreased.
    sqlx = "DECLARE @Aisle CHAR(1) = '" & maisle & "'; DECLARE @Rack VARCHAR(8) = '" & mrack & "';"
    sqlx = sqlx & " DECLARE @NewCapacity INT = (SELECT Capacity FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack);"
    sqlx = sqlx & " DECLARE @OldCapacity INT = (SELECT COUNT(*) FROM RackPos WHERE RackNo IN (SELECT ID FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack));"
    sqlx = sqlx & " IF (@OldCapacity > @NewCapacity) BEGIN"
    sqlx = sqlx & " UPDATE rp SET rp.Posn_Num = s.RowNum FROM RackPos rp INNER JOIN (SELECT ID, ROW_NUMBER() OVER (ORDER BY Count_Qty DESC) AS RowNum"
    sqlx = sqlx & " FROM RackPos WHERE RackNo IN (SELECT ID FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack)) s ON rp.ID = s.ID;"
    ' Then delete the records out of range.
    sqlx = sqlx & " DELETE FROM RackPos WHERE RackNo IN (SELECT ID FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack) AND Posn_Num > @NewCapacity END"
    ' But if capacity has increased, then add the new rows.
    sqlx = sqlx & " ELSE IF (@OldCapacity < @NewCapacity) BEGIN ;WITH cte AS ("
    sqlx = sqlx & " SELECT m.ID + ROW_NUMBER() OVER(ORDER BY rp.ID ASC) AS ID, rp.RackNo, p.Position + ROW_NUMBER() OVER(ORDER BY rp.ID ASC) AS Posn_Num, '' AS SKU, '' AS Lot_Num,"
    sqlx = sqlx & " '' AS Pallet_Num, 0 AS Count_Qty, GETDATE() AS Recv_Date, 'Y' AS BBC, '' AS BarCode, '' AS Lot2, 0 AS Qty2, 'Y' AS Wrapped, 'Y' AS Hold FROM RackPos rp"
    sqlx = sqlx & " CROSS APPLY (SELECT MAX(ID) AS ID FROM RackPos) m"
    sqlx = sqlx & " CROSS APPLY (SELECT MAX(Posn_Num) AS Position FROM RackPos WHERE RackNo IN (SELECT ID FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack)) p"
    sqlx = sqlx & " WHERE RackNo IN (SELECT ID FROM Racks WHERE Aisle = @Aisle AND Rack = @Rack))"
    sqlx = sqlx & " INSERT INTO RackPos (ID, RackNo, Posn_Num, SKU, Lot_Num, Pallet_Num, Count_Qty, Recv_Date, BBC, BarCode, Lot2, Qty2, Wrapped, Hold)"
    sqlx = sqlx & " SELECT ID, RackNo, Posn_Num, SKU, Lot_Num, Pallet_Num, Count_Qty, Recv_Date, BBC, BarCode, Lot2, Qty2, Wrapped, Hold FROM cte WHERE Posn_Num <= @NewCapacity"
    sqlx = sqlx & " UPDATE sequences SET sequence_id = (SELECT MAX(ID) FROM RackPos) WHERE seq = 'RackPos'; END"
    Wdb.Execute sqlx
    y = 0
    For i = 1 To Form6.Targets.Rows - 1
        If Val(Form6.Targets.TextMatrix(i, 0)) = Val(RKey) Then
            y = i: Exit For
        End If
    Next i
    If y > 0 Then
        Form6.Targets.TextMatrix(y, 3) = " " & Text1
        Form6.Targets.TextMatrix(y, 4) = " " & Text2
        Form6.Targets.TextMatrix(y, 5) = Val(Text3)
        Form6.Targets.TextMatrix(y, 6) = Val(Text4)
        Form6.Targets.TextMatrix(y, 7) = " " & Text5
        Form6.Targets.TextMatrix(y, 8) = " " & Text6
        Form6.Targets.TextMatrix(y, 2) = Val(Text7)
        Form6.Targets.TextMatrix(y, 12) = Val(Text8)
        Form6.Targets.TextMatrix(y, 9) = Check1.Value
        Form6.Targets.TextMatrix(y, 10) = Check2.Value
        If Label5 > "  " Then
            Form6.Targets.TextMatrix(y, 11) = " " & Label5
        Else
            Form6.Targets.TextMatrix(y, 11) = " " & Label8
        End If
    End If
    refresh_grid1
    Call Form4.Refresh_racks(Form4.RGrid.Row)
End Sub

Private Sub Command10_Click()
    Dim ds As ADODB.Recordset, s As String, t As String, oc As String, sqlx As String
    If Grid1.Row < 1 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 2) < "100" Then Exit Sub
    oc = InputBox("Operation Code:", "Operation Code...", "500")
    If Len(oc) = 0 Then Exit Sub                        'jv041015
    If Len(oc) > 3 Then Exit Sub                        'jv041015
    If Len(oc) = 1 Then                                 'jv041015
        oc = " " & oc & " "                             'jv041015
    Else                                                'jv041015
        If Len(oc) = 2 Then oc = " " & oc               'jv041015
    End If                                              'jv041015
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 4
    t = Left(Grid1.TextMatrix(Grid1.Row, 10), 13)       'jv040115
    s = "select id, sku, lot_num, barcode from rackpos"
    s = s & " where sku = '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
    s = s & " and lot_num = '" & Grid1.TextMatrix(Grid1.Row, 3) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9) & ds(1) & Chr(9) & ds(2) & Chr(9) & ds(3)
            If Left(ds(3), 13) = t Then                 'jv041015
                sqlx = "Update rackpos set barcode = '" & Left(t, 10) & UCase(oc) & Right(ds(3), 3) & "'"
                sqlx = sqlx & " Where id = " & ds!id
                Wdb.Execute sqlx
                Grid2.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "^ID|^SKU|<Lotnum|<BarCode"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 2000
    DoEvents
    refresh_grid1
End Sub

Private Sub Command11_Click()
    Dim mlot1 As String, mlot2 As String, mqty1 As String, mqty2 As String, p As ptask
    Dim s As String, preas As String, cfile As String
    If Grid1.TextMatrix(Grid1.Row, 12) <> "Y" Then
        MsgBox "Edit is available for on-hold units.", vbOKOnly + vbInformation, "Pallet is not on hold.."
        Exit Sub
    End If
    mlot1 = Grid1.TextMatrix(Grid1.Row, 3)
    mqty1 = Val(Grid1.TextMatrix(Grid1.Row, 5))
    mlot2 = Grid1.TextMatrix(Grid1.Row, 6)
    mqty2 = Val(Grid1.TextMatrix(Grid1.Row, 7))
    If Val(mqty1) > 0 Then
        mqty1 = InputBox("Units:", "Lot " & mlot1 & " units.", mqty1)
        If Len(mqty1) = 0 Then Exit Sub
        If Val(mqty1) <= 0 Then
            MsgBox "Quantity: " & mqty1 & " is invalid.", vbOKOnly + vbInformation, "Sorry, try again.."
            Exit Sub
        End If
    End If
    If Val(mqty2) > 0 Then
        mqty2 = InputBox("Units:", "Lot " & mlot2 & " units.", mqty2)
        If Len(mqty2) = 0 Then Exit Sub
        If Val(mqty2) <= 0 Then
            MsgBox "Quantity: " & mqty2 & " is invalid.", vbOKOnly + vbInformation, "Sorry, try again.."
            Exit Sub
        End If
    End If
    
    s = "Update rackpos set count_qty = " & mqty1 & ", qty2 = " & mqty2
    s = s & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute s
    s = "Update pallets set qty1 = " & mqty1 & ", qty2 = " & mqty2
    s = s & " Where barcode = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
    Wdb.Execute s
    
    preas = InputBox("Reason for edit:", "Reason for edit....")
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    p.area = "WMS"
    If Len(preas) > 0 Then
        p.description = preas
    Else
        p.description = " "
    End If
    p.source = "Edit"
    p.target = Me.RKey
    p.product = Grid1.TextMatrix(Grid1.Row, 2) & " " & UCase(Grid1.TextMatrix(Grid1.Row, 14))
    p.palletid = Grid1.TextMatrix(Grid1.Row, 10)
    p.qty = "1"
    p.uom = "Pallet"
    p.lotnum = mlot1
    p.units = mqty1
    p.lotnum2 = mlot2
    p.units2 = mqty2
    p.status = "COMP"
    p.userid = Form1.userid
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    Write #1, "0";
    Write #1, p.area;
    Write #1, p.description;
    Write #1, p.source;
    Write #1, p.target;
    Write #1, p.product;
    Write #1, p.palletid;
    Write #1, p.qty;
    Write #1, p.uom;
    Write #1, p.lotnum;
    Write #1, p.units;
    Write #1, p.lotnum2;
    Write #1, p.units2;
    Write #1, p.status;
    Write #1, p.userid;
    Write #1, p.trandate;
    Write #1, p.reqid
    Close #1
    
    refresh_grid1
End Sub

Private Sub Command12_Click()
    'ba120517
End Sub

Private Sub Command2_Click()
    Dim ds As ADODB.Recordset, s As String, i As Integer, y As Integer, p As ptask
    Dim r As Long, preas As String
    Dim maisle As String, mrack As String                   'jv062111
    maisle = Left(RKey.Caption, 1)                          'jv062111
    mrack = Right(RKey.Caption, Len(RKey.Caption) - 2)      'jv062111
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Clear all products from rack...") = vbNo Then Exit Sub
    preas = InputBox("Reason for delete:", "Reason for delete....")
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    For y = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(y, 2) > "0" Then
            p.area = "WMS"
            If Len(preas) > 0 Then
                p.description = preas
            Else
                p.description = " "
            End If
            p.source = "Delete"
            p.target = Me.RKey
            p.product = Grid1.TextMatrix(y, 2) & " " & UCase(Grid1.TextMatrix(y, 14))
            p.palletid = Grid1.TextMatrix(y, 10)
            p.qty = "1"
            p.uom = "Pallet"
            p.lotnum = Grid1.TextMatrix(y, 3)
            p.units = Grid1.TextMatrix(y, 5)
            p.lotnum2 = Grid1.TextMatrix(y, 6)
            p.units2 = Grid1.TextMatrix(y, 7)
            p.status = "COMP"
            p.userid = "WMS"
            p.trandate = Format(Now, "yyMMdd hh:mm:ss")
            p.reqid = ".."
            Write #1, y;
            Write #1, p.area;
            Write #1, p.description;
            Write #1, p.source;
            Write #1, p.target;
            Write #1, p.product;
            Write #1, p.palletid;
            Write #1, p.qty;
            Write #1, p.uom;
            Write #1, p.lotnum;
            Write #1, p.units;
            Write #1, p.lotnum2;
            Write #1, p.units2;
            Write #1, p.status;
            Write #1, p.userid;
            Write #1, p.trandate;
            Write #1, p.reqid
        End If
    Next y
    Close #1
    
    r = 0
    s = "select * from racks where aisle = '" & maisle & "'"
    s = s & " and rack = '" & mrack & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        r = ds!id
        sqlx = "Update racks set sku = ' ', lot_num = ' ', qty = 0, qty4 = 0, resv_sku = ' ', resv_lot = ' '"
        sqlx = sqlx & ", fo = 0, hold = 0 Where id = " & r
        Wdb.Execute sqlx
    End If
    ds.Close
        
    sqlx = "Update rackpos set sku = ' ', lot_num = ' ', pallet_num = ' ', count_qty = 0, bbc = 'Y'"
    sqlx = sqlx & ", barcode = ' ', lot2 = ' ', qty2 = 0, wrapped = 'Y', hold = 'N'"
    sqlx = sqlx & " Where rackno = " & r
    Wdb.Execute sqlx
    
    Text1 = "": Text2 = "": Text3 = ""
    Text4 = "": Text5 = "": Text6 = ""
    Check1.Value = 0: Check2.Value = 0
    Command1_Click
    DoEvents
    refresh_grid1
End Sub

Private Sub Command3_Click()        'Clear Position
    Dim ds As ADODB.Recordset, y As Integer
    Dim sqlx As String, i As Integer, pqty As Integer
    Dim olot As String, prec As Long
    Dim rsku As String, rlot As String, preas As String
    Dim p As ptask
    If Grid1.Row < 1 Then Exit Sub
    y = Grid1.Row
    If MsgBox("Ok to clear row " & Grid1.Row & "?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    preas = InputBox("Reason for delete:", "Reason for delete....")
    p.area = "WMS"
    If Len(preas) > 0 Then
        p.description = preas
    Else
        p.description = " "
    End If
    p.source = "Delete"
    p.target = Me.RKey
    p.product = Grid1.TextMatrix(y, 2) & " " & UCase(Grid1.TextMatrix(y, 14))
    p.palletid = Grid1.TextMatrix(y, 10)
    p.qty = "1"
    p.uom = "Pallet"
    p.lotnum = Grid1.TextMatrix(y, 3)
    p.units = Grid1.TextMatrix(y, 5)
    p.lotnum2 = Grid1.TextMatrix(y, 6)
    p.units2 = Grid1.TextMatrix(y, 7)
    p.status = "COMP"
    p.userid = "WMS"
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    If LCase(Form1.userid) <> "jvierus" Then
        Open cfile For Append Shared As #1
        Write #1, y;
        Write #1, p.area;
        Write #1, p.description;
        Write #1, p.source;
        Write #1, p.target;
        Write #1, p.product;
        Write #1, p.palletid;
        Write #1, p.qty;
        Write #1, p.uom;
        Write #1, p.lotnum;
        Write #1, p.units;
        Write #1, p.lotnum2;
        Write #1, p.units2;
        Write #1, p.status;
        Write #1, p.userid;
        Write #1, p.trandate;
        Write #1, p.reqid
        Close #1
    End If
        
    sqlx = "Update rackpos Set sku = ' ', lot_num = ' ', pallet_num = ' ', count_qty = 0"
    sqlx = sqlx & ", recv_date = '" & Format(Now, "M-d-yyyy") & "'"
    sqlx = sqlx & ", bbc = 'Y', barcode = ' ', lot2 = ' ', qty2 = 0, wrapped = 'Y', hold = 'N'"
    sqlx = sqlx & " Where id = " & Grid1.TextMatrix(y, 0)
    Wdb.Execute sqlx
    rsku = Grid1.TextMatrix(y, 2)
    rlot = Grid1.TextMatrix(y, 3)
    Grid1.TextMatrix(y, 2) = " ": Grid1.TextMatrix(y, 3) = " "
    Grid1.TextMatrix(y, 4) = " ": Grid1.TextMatrix(y, 5) = " "
    Grid1.TextMatrix(y, 6) = " ": Grid1.TextMatrix(y, 7) = " "
    Grid1.TextMatrix(y, 8) = Format$(Now, "m-dd-yyyy")
    Grid1.TextMatrix(y, 9) = " "
    Grid1.TextMatrix(y, 10) = " "
    Grid1.TextMatrix(y, 11) = "Y"
    Grid1.TextMatrix(y, 12) = " "
    Grid1.TextMatrix(y, 14) = " "
    
    pqty = 0: pqty4 = 0: prec = Grid1.TextMatrix(y, 13)
    olot = "99999"
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 13)) = prec Then
            If Val(Grid1.TextMatrix(i, 5)) > 0 Then
                If Grid1.TextMatrix(i, 9) > " " Then
                    pqty4 = pqty4 + 1
                Else
                    pqty = pqty + 1
                End If
                If rsku = Grid1.TextMatrix(i, 2) Then
                    If Grid1.TextMatrix(i, 3) < olot Then
                        olot = Grid1.TextMatrix(i, 3)
                    End If
                End If
            End If
        End If
    Next i
    sqlx = "select * from racks where id = " & prec
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If pqty + pqty4 = 0 Then
            sqlx = "Update racks set sku = ' ', lot_num = ' ', qty = 0, qty4 = 0, resv_sku = ' '"
            sqlx = sqlx & ", resv_lot = ' ' Where id = " & ds!id
            Text1.Text = " "
            Text2.Text = " "
            Text3.Text = "0"
            Text4.Text = "0"
            Text5.Text = " "
            Text6.Text = " "
            Check1.Value = 0
            Check2.Value = 0
        Else
            sqlx = "Update racks set qty = " & pqty & ", qty4 = " & pqty4 & ", lot_num = '" & olot & "'"
            sqlx = sqlx & " Where id = " & ds!id
            Text3.Text = pqty
            Text4.Text = pqty4
            Text9.Text = pqty
            Text10.Text = pqty4
            Text2.Text = olot
        End If
        Wdb.Execute sqlx
    Else
        MsgBox "Problems with this rack..", vbOKOnly + vbExclamation, "Cannot update.."
    End If
    ds.Close
    If y < Grid1.Rows - 1 Then y = y + 1
    Grid1.Col = 1: Grid1.Row = y
    Call Command1_Click
    Call Grid1_Click
End Sub

Private Sub Command4_Click()            'Insert Pallet
    Dim psku As String, ppal As String, pdesc As String
    Dim plot As String, pdate As String, sqlx As String
    Dim olot As String, psize As String, pbar As String, popl As String
    Dim i As Integer, pqty As Integer, lqty As Integer
    Dim pqty2 As Integer, plot2 As String, p As ptask
    Dim ds As ADODB.Recordset, y As Integer
    Dim pplate As String                                                  'jv070314
    psku = " ": ppal = "0": psize = "0": popl = "_"
    If Val(Grid1.TextMatrix(Grid1.Row, 5)) > 0 Then
        MsgBox "Position currently contains a pallet.", vbOKOnly, "Cannot Insert Here..."
        Exit Sub
    End If
    y = Grid1.Row: lqty = 0: lqty4 = 0
    olot = "99999"
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 5)) > 0 Then
            If Grid1.TextMatrix(i, 9) > " " Then
                lqty4 = lqty4 + 1
            Else
                lqty = lqty + 1
            End If
            psku = Grid1.TextMatrix(i, 2)
            plot = Grid1.TextMatrix(i, 3)
            If plot < olot Then olot = plot
            If Val(Grid1.TextMatrix(i, 4)) >= Val(ppal) Then
                ppal = Val(Grid1.TextMatrix(i, 4)) + 1
            End If
            pqty = Val(Grid1.TextMatrix(i, 5))
            pdate = Grid1.TextMatrix(i, 8)
            'popl = Mid(Grid1.TextMatrix(i, 10), 12, 1)
            popl = Mid(Grid1.TextMatrix(i, 10), 11, 3)                  'jv052515
        End If
    Next i
    
    'User Prompts
    psku = InputBox("SKU #", "Insert Position " & y, psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot #", "Insert Position " & y, plot)
    If Len(plot) = 0 Then Exit Sub
    ppal = InputBox("Pallet #", "Insert Position " & y, ppal)
    If Len(ppal) = 0 Then Exit Sub
    popl = InputBox("Operation Code", "Insert Position " & y, popl)
    'If Len(popl) = 0 Or Len(popl) > 1 Then Exit Sub
    If Len(popl) = 0 Or Len(popl) > 3 Then Exit Sub                     'jv052515
    If Len(popl) = 1 Then                                               'jv052515
        popl = " " & popl & " "                                         'jv052515
    Else                                                                'jv052515
        If Len(popl) = 2 Then popl = " " & popl                         'jv052515
    End If                                                              'jv052515
    psize = InputBox("4Way Size:", "Insert Position " & y, 0)
    If Len(psize) = 0 Then Exit Sub
    psku = Trim(Left(psku, 4))                                          'jv082415
    plot = Left$(plot, 5)
    ppal = UCase(ppal)
    popl = UCase(popl)
    If skurec(Val(psku)).sku <> psku Then
        pdesc = "Ingredient"
        If Val(psize) = 0 Then
            pqty = 1
            lqty = lqty + 1
        Else
            pqty = CInt(psize)
            lqty4 = lqty4 + 1
        End If
    Else
        pdesc = skurec(Val(psku)).prodname
        If Val(psize) = 0 Then
            pqty = skurec(Val(psku)).uom_per_pallet
            lqty = lqty + 1
        Else
            pqty = CInt(psize)
            lqty4 = lqty4 + 1
        End If
    End If
    pdate = Format$(Now, "m-d-yyyy")
    pqty2 = pqty
    pqty2 = InputBox("Unit Qty for Lot " & plot & ":", "Lot " & plot & " units...", pqty2)
    If Len(pqty2) = 0 Then Exit Sub
    If pqty2 <> pqty Then
        i = pqty - pqty2    '308 - 200 = 108
        pqty = pqty2        '200
        pqty2 = i           '108
        plot2 = Format(Val(plot) + 1, "00000") & popl                                          'jv052515
        plot2 = InputBox("Lot 2#:", "Insert Position " & y, plot2)
        If Len(plot2) = 0 Then
            plot2 = " "
            pqty2 = 0
        End If
    Else
        plot2 = " "
        pqty2 = 0
    End If
    
    sqlx = "select * from rackpos"
    sqlx = sqlx & " where id = " & Grid1.TextMatrix(y, 0)
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update rackpos set sku = '" & psku & "'"
        If Mid(Me.Caption, 6, 1) = "M" Then
            sqlx = sqlx & ", posn_num = " & order_pick_position(psku)
            Grid1.TextMatrix(y, 1) = order_pick_position(psku)
        End If
        sqlx = sqlx & ", lot_num = '" & plot & "'"
        sqlx = sqlx & ", pallet_num = '" & ppal & "'"
        sqlx = sqlx & ", count_qty = " & pqty
        sqlx = sqlx & ", recv_date = '" & Format(Now, "M-d-yyyy") & "'"
        If psize > 0 Then
            sqlx = sqlx & ", bbc = 'N'"
        Else
            sqlx = sqlx & ", bbc = 'Y'"
        End If
        pbar = psku
        If Len(psku) = 3 Then                                               'jv082415
            pbar = pbar & " " & Form1.bb_codedate(plot)                     'jv082415
        Else                                                                'jv082415
            pbar = pbar & Form1.bb_codedate(plot)                           'jv082415
        End If                                                              'jv082415
        pbar = pbar & popl & Format(ppal, "000")                                'jv052515
        sqlx = sqlx & ", barcode = '" & pbar & "'"
        sqlx = sqlx & ", lot2 = '" & plot2 & "'"
        sqlx = sqlx & ", qty2 = " & pqty2
        sqlx = sqlx & ", wrapped = 'Y'"
        sqlx = sqlx & ", hold = 'N'"
        sqlx = sqlx & " where id = " & Grid1.TextMatrix(y, 0)
        Wdb.Execute sqlx
    End If
    ds.Close
    Grid1.TextMatrix(y, 2) = psku
    Grid1.TextMatrix(y, 3) = plot
    Grid1.TextMatrix(y, 4) = ppal
    Grid1.TextMatrix(y, 5) = pqty
    Grid1.TextMatrix(y, 6) = plot2
    Grid1.TextMatrix(y, 7) = pqty2
    Grid1.TextMatrix(y, 8) = pdate
    If psize > 0 Then Grid1.TextMatrix(y, 9) = "Y"
    Grid1.TextMatrix(y, 10) = pbar
    Grid1.TextMatrix(y, 11) = "Y"
    Grid1.TextMatrix(y, 12) = " "
    Grid1.TextMatrix(y, 14) = StrConv(pdesc, vbProperCase)
    If plot < olot Then olot = plot
    p.area = "WMS"
    p.description = " "
    p.source = "Insert"
    p.target = Me.RKey
    p.product = Grid1.TextMatrix(y, 2) & " " & UCase(Grid1.TextMatrix(y, 14))
    p.palletid = Grid1.TextMatrix(y, 10)
    p.qty = "1"
    p.uom = "Pallet"
    p.lotnum = Grid1.TextMatrix(y, 3)
    p.units = Grid1.TextMatrix(y, 5)
    p.lotnum2 = Grid1.TextMatrix(y, 6)
    p.units2 = Grid1.TextMatrix(y, 7)
    p.status = "COMP"
    p.userid = "WMS"
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    Write #1, y;
    Write #1, p.area;
    Write #1, p.description;
    Write #1, p.source;
    Write #1, p.target;
    Write #1, p.product;
    Write #1, p.palletid;
    Write #1, p.qty;
    Write #1, p.uom;
    Write #1, p.lotnum;
    Write #1, p.units;
    Write #1, p.lotnum2;
    Write #1, p.units2;
    Write #1, p.status;
    Write #1, p.userid;
    Write #1, p.trandate;
    Write #1, p.reqid
    Close #1
    
    sqlx = "Update racks set qty = " & lqty & ", qty4 = " & lqty4 & ", sku = '" & psku & "'"
    sqlx = sqlx & ", lot_num = '" & olot & "' Where id = " & Grid1.TextMatrix(y, 13)
    Wdb.Execute sqlx
    
    'Update pallet record
    If Val(psku) >= 100 Then                                                    'jv070314
        recid = 0
        s = "select * from pallets where barcode = '" & pbar & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            pplate = ds!plateno                                                 'jv070314
            recid = ds!id
        Else
            ds.Close
            s = "select * from pallets where status in ('Shipped','Order Pick')"
            s = s & " order by trandate"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                pplate = " "                                                    'jv070314
                recid = ds!id
            End If
        End If
        ds.Close
        If recid > 0 Then
            s = "Update pallets set plateno = '" & pplate & "'"                 'jv070314
            s = s & ",barcode = '" & pbar & "'"
            s = s & ",qty1 = " & Val(pqty)
            s = s & ",lot1 = '" & plot & "'"
            s = s & ",qty2 = " & Val(pqty2)
            s = s & ",lot2 = '" & plot2 & "'"
            s = s & ",source = 'Racks'"
            s = s & ",target = '" & RKey.Caption & "'"
            If psize = 0 Then
                s = s & ",bbc = 'Y'"
            Else
                s = s & ",bbc = 'N'"
            End If
            If RKey = "M-OP" Then
                s = s & ",status = 'Order Pick'"
            Else
                s = s & ",status = 'Warehouse'"
            End If
            s = s & ",trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",sku = '" & psku & "'"
            s = s & " Where id = " & recid
            Wdb.Execute s
        Else
            pid = wd_seq("Pallets")
            s = "Insert Into pallets Values (" & pid
            s = s & ",'" & pplate & "'"                                         'jv070314
            s = s & ",'" & pbar & "'"
            s = s & "," & Val(pqty)
            s = s & ",'" & plot & "'"
            s = s & "," & Val(pqty2)
            s = s & ",'" & plot2 & "'"
            s = s & ",'Racks'"
            s = s & ",'" & RKey.Caption & "'"
            If psize = 0 Then
                s = s & ",'Y'"
            Else
                s = s & ",'N'"
            End If
            If RKey = "M-OP" Then
                s = s & ",'Order Pick'"
            Else
                s = s & ",'Warehouse'"
            End If
            s = s & ",'" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",'" & psku & "')"
            Wdb.Execute s
        End If
    End If                                                                              'jv070314
        
    Text3.Text = lqty
    Text4.Text = lqty4
    Text9.Text = lqty
    Text10.Text = lqty4
    Text1.Text = psku
    Text2.Text = olot
    If y > 1 Then y = y - 1
    Grid1.Col = 1: Grid1.Row = y
    Call Command1_Click
    Call Grid1_Click
End Sub

Private Sub Command5_Click()
    Dim s As String, p As String, rid As Long
    p = Val(Grid1.TextMatrix(Grid1.Row, 1))
    p = InputBox("Posn Num:", "Posn Number...", p)
    If Len(p) = 0 Then Exit Sub
    rid = wd_seq("RackPos")
    s = "INSERT INTO RackPos (ID, RackNo, Posn_Num, SKU, Lot_Num, Pallet_Num,"
    s = s & " Count_Qty, Recv_Date, BBC, BarCode, Lot2, Qty2, Wrapped, Hold)"
    s = s & " VALUES (" & rid & ","
    s = s & RKey.Caption & ","
    s = s & Val(p) & ","
    s = s & "' ',"
    s = s & "'.',"
    s = s & "'.',"
    s = s & "0,"
    s = s & "'" & Format(Now, "mm-dd-yyyy") & "',"
    s = s & "'Y',"
    s = s & "'.',"
    s = s & "'.',"
    s = s & "0,"
    s = s & "'Y',"
    s = s & "'N')"
    Wdb.Execute s
    refresh_grid1
End Sub

Private Sub Command6_Click()
    Dim sqlx As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 11) = "Y" Then
        Grid1.TextMatrix(Grid1.Row, 11) = " "
    Else
        Grid1.TextMatrix(Grid1.Row, 11) = "Y"
    End If
    If Grid1.TextMatrix(Grid1.Row, 11) = "Y" Then
        sqlx = "Update rackpos set wrapped = 'Y' Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Else
        sqlx = "Update rackpos set wrapped = 'N' Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    End If
    Wdb.Execute sqlx
    
    Grid1.FillStyle = flexFillRepeat
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    If Grid1.TextMatrix(Grid1.Row, 10) > "0" And Grid1.TextMatrix(Grid1.Row, 11) <> "Y" Then
        Grid1.CellBackColor = RKey.BackColor
    Else
        Grid1.CellBackColor = Grid1.BackColor
    End If
End Sub

Private Sub Command7_Click()
    Dim psku As String, plot As String, pcode As String, ppal As String, zid As Long
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    psku = Grid1.TextMatrix(Grid1.Row, 2)                                                       'jv040915
    plot = Grid1.TextMatrix(Grid1.Row, 3)                                                       'jv040915
    If psku < "000" Or psku > "9999" Then Exit Sub                                               'jv040915
    If plot < "00000" Or plot > "99999" Then Exit Sub                                           'jv040915
    If Grid1.TextMatrix(Grid1.Row, 12) = "Y" Then
        Grid1.TextMatrix(Grid1.Row, 12) = " "
    Else
        Grid1.TextMatrix(Grid1.Row, 12) = "Y"
    End If
    zid = wd_seq("HoldList")                            'jv042015
    psku = Grid1.TextMatrix(Grid1.Row, 2)                                                       'jv040715
    plot = Grid1.TextMatrix(Grid1.Row, 3)                                                       'jv040715
    pcode = Trim(Mid(Grid1.TextMatrix(Grid1.Row, 10), 11, 3))                              'jv052515
    ppal = Mid(Grid1.TextMatrix(Grid1.Row, 10), 14, 3)                                          'jv040715
    If Grid1.TextMatrix(Grid1.Row, 12) = "Y" Then                                               'jv040715
        s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate) values (" & zid  'jv040715
        s = s & ", '" & psku & "', '" & plot & "', '" & pcode & "', '" & ppal & "', '" & ppal & "', 'Racks', '" & WDUserId & "'"
        s = s & ", '" & Format(Now, "yyMMdd hh:mm:ss") & "')"
        Wdb.Execute s                                                        'jv040715
    Else                                                                    'jv040715
        s = "delete from holdlist where sku = '" & psku & "'"               'jv040715
        s = s & " and lot_num = '" & plot & "'"                             'jv040715
        s = s & " and opcode = '" & pcode & "'"                             'jv040715
        s = s & " and spallet = '" & ppal & "'"                             'jv040715
        s = s & " and epallet = '" & ppal & "'"                             'jv040715
        Wdb.Execute s                                                        'jv040715
    End If                                                                  'jv040715
        
    Grid1.FillStyle = flexFillRepeat
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    If Grid1.TextMatrix(Grid1.Row, 10) > "0" And Grid1.TextMatrix(Grid1.Row, 11) <> "Y" Then
        Grid1.CellBackColor = RKey.BackColor
    Else
        Grid1.CellBackColor = Grid1.BackColor
    End If
End Sub

Private Sub Command8_Click()
    Dim i As Integer, rno As Long, s As String
    Dim pqty As Integer, pqty4 As Integer
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    rno = Val(Grid1.TextMatrix(Grid1.Row, 13))
    For i = 1 To Grid1.Rows - 1
        s = "update rackpos set rackno = " & rno
        s = s & " where id = " & Grid1.TextMatrix(i, 0)
        Wdb.Execute s
        If rno <> Grid1.TextMatrix(i, 13) Then
            s = "delete from racks where id = " & Grid1.TextMatrix(i, 13)
            Wdb.Execute s
        End If
    Next i
    pqty = 0: pqty4 = 0
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 2) > "0" Then
            If Grid1.TextMatrix(i, 8) = "Y" Then
                pqty4 = pqty4 + 1
            Else
                pqty = pqty + 1
            End If
        End If
    Next i
    s = "update racks set capacity=" & Grid1.Rows - 1
    s = s & ", qty=" & pqty
    s = s & ", qty4=" & pqty4
    s = s & " where id = " & rno
    Wdb.Execute s
    refresh_grid1
End Sub

Private Sub Command9_Click()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String
    Dim cfile As String, s As String, bc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String
    Dim logpath As String
    slit.Caption = " "
    logpath = Form1.logdir
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) < 1 Then Exit Sub
    sdate = Format(Grid1.TextMatrix(Grid1.Row, 8), "yyyymmdd")
    sdate = InputBox("Start Date (YearMoDa):", "Start Date...", sdate)
    If Len(sdate) = 0 Then Exit Sub
    edate = InputBox("End Date (YearMoDa):", "End Date...", Format(Now, "yyyymmdd"))
    If Len(edate) = 0 Then Exit Sub
    
    wbc = Grid1.TextMatrix(Grid1.Row, 10)
    wbc = InputBox("Enter a BarCode to search for:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub

    Screen.MousePointer = 11
    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    s = f2 & " " & f4 & " " & f5
                    Screen.MousePointer = 0
                    'MsgBox s, vbOKOnly + vbInformation, f15 & " shipped...... " & f6
                    slit.Caption = s & " " & f15 & " shipped...... " & f6
                    Close #1
                    Command3_Click
                    slit.Caption = " "
                    Exit Sub
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    If RKey = "M-OP" Or RKey = "M-SP" Then
        spath = logpath & "move*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        s = f2 & " " & f4 & " " & f5
                        Screen.MousePointer = 0
                        'MsgBox s, vbOKOnly + vbInformation, f15 & " shipped...... " & f6
                        slit.Caption = s & " " & f15 & " shipped...... " & f6
                        Close #1
                        Command3_Click
                        slit.Caption = " "
                        Exit Sub
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form5.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form5" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form5.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form5.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form5.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form5.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Form5.ActiveControl.Name = "Check2" Then
            Text1.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    End If
    If KeyAscii = 27 Then
        Call Form_Deactivate
        Unload Form5
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        Call Command1_Click 'F2-Accept
        Call Form_Deactivate
        Unload Form5
    End If
    If KeyCode = 120 Then Call Command2_Click 'F9-Clear
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form5" Then
            Form5.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form5.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form5.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form5.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    possku.Caption = "..."
    posdesc.Caption = "..."
    If Form1.plantno <> "50" Then
        Command9.Visible = False
        slit.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_Click()
    Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu viewmenu
End Sub

Private Sub Grid1_RowColChange()
    If Len(Grid1.TextMatrix(Grid1.Row, 2)) > 0 Then
        possku.Caption = Grid1.TextMatrix(Grid1.Row, 2)
    Else
        possku.Caption = "..."
        posdesc.Caption = "..."
    End If
End Sub

Private Sub palhist_Click()
    palhistory.Show
    palhistory.barkey = Grid1.TextMatrix(Grid1.Row, 10)
End Sub

Private Sub possku_Change()
    If possku = "..." Then Exit Sub
    posdesc.Caption = "..."
    posdesc.Caption = skurec(Val(possku)).prodname
End Sub

Private Sub RKey_Change()
    DoEvents
    refresh_grid1
End Sub

Private Sub Text1_Change()
    Dim i As Integer
    i = Val(Trim(Text1))
    Label5 = skurec(i).prodname
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Val(Text3) = 0 Then Text3 = ""
End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Val(Text4) = 0 Then Text4 = ""
End Sub

Private Sub Text5_Change()
    Dim i As Integer
    i = Val(Trim(Text5))
    Label8 = skurec(i).prodname
End Sub

Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text6_GotFocus()
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Text7_GotFocus()
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.Text)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If Val(Text7) = 0 Then Text7 = ""
End Sub

Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If Val(Text8) = 0 Then Text8 = ""
End Sub
