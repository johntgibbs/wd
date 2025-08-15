VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Rastrail 
   Caption         =   "Post Trailers To Trailer History"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form2"
   ScaleHeight     =   4830
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   7080
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid hg 
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5741
      _Version        =   327680
      Cols            =   8
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Ship Date"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Plant"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Rastrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_lists()
    Dim ds As adodb.Recordset, s As String
    Combo1.Clear: Combo2.Clear
    s = "select * from plants"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!plantname
            List1.AddItem ds!plant
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select distinct shipdate from trailers"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo2.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
End Sub

Private Sub postsheet()
    Dim ds As adodb.Recordset
    Dim sqlx As String, i As Integer, k As Integer, z As Long
    hg.Visible = False
    hg.Clear: hg.Rows = 100: hg.Cols = 8
    sqlx = "select * from trailers where plant = " & List1
    sqlx = sqlx & " and shipdate <= '" & Combo2 & "'"
    sqlx = sqlx & " and pb_flag = 'Y'"
    sqlx = sqlx & " order by runid, sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            hg.TextMatrix(ds!branch, 0) = ds!branch
            k = 0
            If Left(ds!trlno, 1) = "#" Then k = Val(Right(ds!trlno, 1))
            If Left(ds!trlno, 1) = "B" Then k = 7
            If k = 0 Then k = 6
            If k > 7 Then k = 7
            hg.TextMatrix(ds!branch, k) = Val(hg.TextMatrix(ds!branch, k)) + ds!units
            ds.MoveNext
        Loop
        For i = hg.Rows - 1 To 1 Step -1
            If Val(hg.TextMatrix(i, 0)) = 0 Then
                hg.RemoveItem i
            Else
                z = wd_seq("TrHist", Form1.shipdb)
                s = "Insert into trhist (id, plant, branch, shipdate, trl1, trl2, trl3, trl4, trl5, trl6, trladj, trlbob)"  'jv120315
                s = s & " Values (" & z & ", "
                s = s & Val(List1) & ", "
                s = s & Val(hg.TextMatrix(i, 0)) & ", '"
                s = s & Format(Combo2, "m-d-yyyy") & "', "
                s = s & Val(hg.TextMatrix(i, 1)) & ", "
                s = s & Val(hg.TextMatrix(i, 2)) & ", "
                s = s & Val(hg.TextMatrix(i, 3)) & ", "
                s = s & Val(hg.TextMatrix(i, 4)) & ", "
                s = s & Val(hg.TextMatrix(i, 5)) & ", "
                s = s & Val(hg.TextMatrix(i, 6)) & ", "
                s = s & "0, "                                                   'jv120315
                s = s & Val(hg.TextMatrix(i, 7)) & ")"
                Sdb.Execute s
            End If
        Next i
    End If
    ds.Close
    hg.FormatString = "^Branch|^#1|^#2|^#3|^#4|^#5|^#6|^Bob"
    hg.ColWidth(0) = 800: hg.ColWidth(1) = 800
    hg.ColWidth(2) = 800: hg.ColWidth(3) = 800
    hg.ColWidth(4) = 800: hg.ColWidth(5) = 800
    hg.ColWidth(6) = 800: hg.ColWidth(7) = 800
    hg.Visible = True
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Call postsheet
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Rastrail.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "rastrail" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Rastrail.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Rastrail.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Rastrail.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Rastrail.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "rastrail" Then
            Form1.FrmGrid.Col = 1: Rastrail.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Rastrail.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Rastrail.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Rastrail.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    hg.Row = 0
    hg.FormatString = "^TRHist|^#1|^#2|^#3|^#4|^#5|^#6|^Bob"
    hg.ColWidth(0) = 1000: hg.ColWidth(7) = 800
    hg.ColWidth(1) = 800: hg.ColWidth(2) = 800: hg.ColWidth(3) = 800
    hg.ColWidth(4) = 800: hg.ColWidth(5) = 800: hg.ColWidth(6) = 800
    refresh_lists
End Sub

Private Sub Form_Resize()
    If Rastrail.Height > 2895 Then
        hg.Height = Rastrail.Height - 2000
    Else
        hg.Height = 2895
    End If
    hg.Width = Rastrail.Width - 85
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub
