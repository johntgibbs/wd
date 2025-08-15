VERSION 5.00
Begin VB.Form bimptstimp 
   Caption         =   "Import into Test DB"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Process Trailers"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Runs"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.ListBox List2 
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
      Height          =   4740
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   11415
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
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11415
   End
End
Attribute VB_Name = "bimptstimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_list1()
    Dim ds As ADODB.Recordset, s As String
    List1.Clear
    List1.AddItem "delete from runs"
    s = "select * from runs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate,"
            s = s & " startime, pickup, oc, yardnote, loadnote) values ("
            s = s & ds!id
            s = s & ", '" & ds!loaded & "'"
            s = s & ", '" & ds!destination & "'"
            s = s & ", '" & ds!locname & "'"
            s = s & ", '" & ds!trlno & "'"
            s = s & ", " & ds!trlsize
            s = s & ", '" & Format(ds!trldate, "M-dd-yyyy") & "'"
            s = s & ", '" & Format(ds!startime, "h:mm am/pm") & "'"
            s = s & ", '" & ds!pickup & "'"
            s = s & ", '" & ds!oc & "'"
            s = s & ", '" & ds!yardnote & "'"
            s = s & ", '" & ds!loadnote & "')"
            List1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub refresh_list2()
    Dim ds As ADODB.Recordset, s As String
    List2.Clear
    List2.AddItem "delete from trailers"
    s = "select * from trailers"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate,"
            s = s & " trlno, sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) values ("
            s = s & ds!id
            s = s & ", " & ds!runid
            s = s & ", '" & ds!groupcode & "'"
            s = s & ", " & ds!plant
            s = s & ", " & ds!branch
            s = s & ", '" & ds!account & "'"
            s = s & ", '" & Format(ds!shipdate, "M-dd-yyyy") & "'"
            s = s & ", '" & ds!trlno & "'"
            s = s & ", '" & ds!sku & "'"
            s = s & ", " & ds!pallets
            s = s & ", " & ds!wraps
            s = s & ", " & ds!units
            s = s & ", " & ds!whs_num
            s = s & ", '" & ds!pb_flag & "'"
            s = s & ", '" & ds!ra_flag & "')"
            List2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub Command1_Click()
    Dim tb As ADODB.Connection, i As Integer
    Screen.MousePointer = 11
    Set tb = CreateObject("ADODB.Connection")
    tb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    For i = 0 To List1.ListCount - 1
        tb.Execute List1.List(i)
    Next i
    tb.Close
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    Dim tb As ADODB.Connection, i As Integer
    Screen.MousePointer = 11
    Set tb = CreateObject("ADODB.Connection")
    tb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    For i = 0 To List2.ListCount - 1
        tb.Execute List2.List(i)
    Next i
    tb.Close
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    refresh_list1
    refresh_list2
End Sub

Private Sub Form_Resize()
    List1.Width = Me.Width - 120
    List2.Width = Me.Width - 120
End Sub
