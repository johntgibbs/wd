VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "WMS Plate Application"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   7275
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Wrappers  "
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
      Left            =   360
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.OptionButton Option5 
         Caption         =   "Backhaul"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tri - Level"
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Roller Bed"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Robot Zero"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Snack Plant"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5520
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2520
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   7935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Label Label3 
      Caption         =   "Scanned Labels:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label rcolor 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label emess 
      Alignment       =   2  'Center
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "Plate Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Label BarCode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function check_dai_plate(bc As String) As String
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select plateno from pallets where barcode = '" & bc & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        check_dai_plate = ds!plateno
    Else
        check_dai_plate = "None"
    End If
    ds.Close
    db.Close
End Function

Private Sub record_pallet(pno As String, p As ptask, pwhs As String, pstat As String)
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim pid As Long, psku As String, recid As Long
    Screen.MousePointer = 11
    'pstat = "Wrapper"
    'pstat = "Warehouse"
    'pstat = "Test"
    psku = Trim(Left(p.palletid, 4))
    recid = 0
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    db.Open Form1.bbsr
    
    s = "select * from pallets where barcode = '" & p.palletid & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        recid = ds!id
    Else
        ds.Close
        s = "select * from pallets where status in ('Shipped','Order Pick')"
        s = s & " order by trandate"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            recid = ds!id
        End If
    End If
    ds.Close
    If recid > 0 Then
        s = "Update pallets set plateno = '" & Trim(pno) & "'"
        s = s & ",barcode = '" & p.palletid & "'"
        s = s & ",qty1 = " & Val(p.units)
        s = s & ",lot1 = '" & p.lotnum & "'"
        s = s & ",qty2 = " & Val(p.units2)
        s = s & ",lot2 = '" & p.lotnum2 & "'"
        s = s & ",source = '" & p.source & "'"
        's = s & ",target = '" & Grid1.TextMatrix(i, 4) & "'"
        's = s & ",target = 'SR" & pwhs & "'"
        s = s & ",target = '" & pwhs & "'"
        s = s & ",bbc = 'Y'"
        s = s & ",status = '" & pstat & "'"
        s = s & ",trandate = '" & p.trandate & "'"
        s = s & ",sku = '" & psku & "'"
        s = s & " Where id = " & recid
        'MsgBox s
        db.Execute s
    Else
        pid = wd_seq("Pallets")
        s = "Insert Into pallets Values (" & pid
        s = s & ",'" & Trim(pno) & "'"
        s = s & ",'" & p.palletid & "'"
        s = s & "," & Val(p.units)
        s = s & ",'" & p.lotnum & "'"
        s = s & "," & Val(p.units2)
        s = s & ",'" & p.lotnum2 & "'"
        s = s & ",'" & p.source & "'"
        's = s & ",'" & Grid1.TextMatrix(i, 4) & "'"
        's = s & ",'SR" & pwhs & "'"
        s = s & ",'" & pwhs & "'"
        s = s & ",'Y'"
        If p.target = "ORDER PICK" Then
            s = s & ",'Order Pick'"
        Else
            s = s & ",'" & pstat & "'"
        End If
        s = s & ",'" & p.trandate & "'"
        s = s & ",'" & psku & "')"
        db.Execute s
    End If
    db.Close
    Screen.MousePointer = 0
End Sub

Private Sub refresh_labels()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim i As Long, p As ptask
    Combo1.Clear
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    If Option1 = True Then          'Snack Plant
        s = "select * from paltasks"
        s = s & " where source in ('1405','1406','1731','SNACK PLANT')"
        s = s & " and area = 'DOCK'"
        s = s & " and status <> 'COMP' and userid < '0'"
        s = s & " order by palletid"
    End If
    If Option2 = True Then          'Robot Zero
        s = "select * from paltasks"
        s = s & " where source = 'ROBOT ZERO'"
        s = s & " and area = 'FORKLIFT'"
        s = s & " and status <> 'COMP' and userid < '0'"
        s = s & " order by palletid"
    End If
    If Option3 = True Then          'Roller Bed
        s = "select * from paltasks"
        s = s & " where area = 'ROLLER BED'"
        's = s & " and area = 'FORKLIFT'"
        s = s & " and status <> 'COMP'" ' and userid < '0'"
        s = s & " order by palletid"
    End If
    If Option4 = True Then          'TRI Level
        s = "select * from paltasks"
        s = s & " where area = 'TRAFFIC MASTER'"
        's = s & " and area = 'FORKLIFT'"
        s = s & " and status <> 'COMP' and userid < '0'"
        s = s & " order by palletid"
    End If
    If Option5 = True Then          'Back Haul
        s = "select * from paltasks"
        s = s & " where (source = 'BACKHAUL'"
        s = s & " and area = 'DOCK')"
        s = s & " OR (source = 'STAGING' and area = 'FORKLIFT')"
        s = s & " and status <> 'COMP'" ' and userid < '0'"
        s = s & " order by palletid"
    End If
    
    
    'MsgBox s
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If check_dai_plate(ds!palletid) = "None" Then
                i = wd_seq("BHBarcode")
                p = masterec(ds!id)
                Call record_pallet(Str(i), p, p.target, "Warehouse")
                'MsgBox p.palletid
                DoEvents
            End If
            Combo1.AddItem ds!palletid & "  " & Format(ds!reqid, "000000")
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    If Combo1.ListCount < 1 Then Combo1.AddItem ".."
    'Combo1.ListIndex = 0
End Sub
Private Sub update_plate()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Text1 = UCase(Text1)
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select barcode from pallets where plateno = '" & Text2 & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        s = "Plate number has already been assigned to " & ds!barcode & "."
        'MsgBox s, vbExclamation + vbOKOnly, "try again..."
        emess.BackColor = rcolor.BackColor: emess.ForeColor = rcolor.ForeColor
        emess = "Plate " & Text2 & " already assigned to " & ds!barcode & "."
        ds.Close: db.Close
        Text1 = "": Text2 = ""
        Exit Sub
    End If
    s = "select * from pallets where barcode = '" & Text1 & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update pallets set plateno = '" & Text2 & "'"
        s = s & " Where id = " & ds!id
        db.Execute (s)
        s = "Update paltasks set reqid = '" & Text2 & "'"
        s = s & " Where palletid = '" & Text1 & "'"
        'MsgBox s
        db.Execute (s)
        emess.BackColor = Me.BackColor: emess.ForeColor = Me.ForeColor
        emess = "Plate " & Text2 & " has been assigned to " & Text1 & "."
    Else
        s = "Pallet label has not been scanned at the wrapper."
        'MsgBox s, vbExclamation + vbOKOnly, "try again....."
        emess.BackColor = rcolor.BackColor: emess.ForeColor = rcolor.ForeColor
        emess = "Label " & Text1 & " has not been scanned at the wrapper."
    End If
    ds.Close: db.Close
    Text1 = "": Text2 = ""
End Sub

Private Sub Combo1_Click()
    Text1 = Combo1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text1 = Left(Combo1, 16)
        Text2.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    If Text2 <= " " Then Exit Sub
    update_plate
    DoEvents
    Text1.SetFocus
    refresh_labels
End Sub

Private Sub Command2_Click()
    refresh_labels
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    'bbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    'tbbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    'bbsr = "odbc;database=wdracks;dsn=wdracks"
    'tbbsr = "odbc;database=wdracks;dsn=wdracks"
    'Option4.Value = True
    Label4 = Form1.bbsr
    Text1 = "": Text2 = ""
    refresh_labels
End Sub

Private Sub Form_Resize()
    Label1.Left = (Me.Width - Text1.Width) * 0.5
    Text1.Left = Label1.Left
    Label2.Left = Label1.Left
    Text2.Left = Label1.Left
    emess.Left = Label1.Left
    Label3.Left = Label1.Left
    Combo1.Left = Label1.Left
    Command1.Left = Label1.Left '(Me.Width - Command1.Width) * 0.5
    Command2.Left = Text1.Left + Text1.Width - Command2.Width
End Sub

Private Sub Option1_Click()
    refresh_labels
End Sub

Private Sub Option2_Click()
    refresh_labels
End Sub

Private Sub Option3_Click()
    refresh_labels
End Sub

Private Sub Option4_Click()
    refresh_labels
End Sub

Private Sub Option5_Click()
    refresh_labels
End Sub

Private Sub Text1_Change()
    If Len(Text1) > 15 Then
        Text1 = UCase(Text1)
        Text2.SetFocus
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If emess.BackColor = rcolor.BackColor Then
        emess.BackColor = Me.BackColor: emess.ForeColor = Me.ForeColor
        emess = "..."
    End If
End Sub

Private Sub Text2_Change()
    If Len(Text2) > 5 Then Command1.SetFocus
End Sub
