VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Prodsources 
   Caption         =   "Production Source Listing"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2566
      _Version        =   327680
      BackColorFixed  =   65535
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "Prodsources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select * from prodsources order by source"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!source & Chr(9)
            s = s & ds!sourcename & Chr(9)
            If ds!tl_flag = "Y" Then s = s & "*"
            s = s & Chr(9)
            s = s & ds!days
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Source|<Name|^Tri-level|^Days"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 2500
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
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

Private Sub Form_Deactivate()
    Dim i As Integer
    If Prodsources.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "prodsources" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Prodsources.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Prodsources.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Prodsources.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Prodsources.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "prodsources" Then
            Form1.FrmGrid.Col = 1: Prodsources.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Prodsources.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Prodsources.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Prodsources.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 10: Grid1.FontBold = True
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 380 '680
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub
