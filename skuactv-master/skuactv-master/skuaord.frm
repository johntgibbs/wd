VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Pallet Orders"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form5"
   ScaleHeight     =   3135
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      _Version        =   327680
      Cols            =   7
      BackColorFixed  =   12648447
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label oprod 
      Caption         =   "oprod"
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
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label osku 
      Caption         =   "osku"
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
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, psz As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1
    If Form1.Check1.Value = 1 Then
        psz = "??"
        s = "select uom_per_pallet from sku_config where sku = '" & osku & "'"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            psz = ds!uom_per_pallet
        End If
        ds.Close
        s = "select * from ship_infc where sku = '" & osku & "'"
        s = s & " and ship_status not in ('CANC','DONE')"
        s = s & " order by ship_date,order_num"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!order_num & Chr(9)
                s = s & ds!to_whse_num & Chr(9)
                s = s & Format(ds!ship_date, "m-d-yyyy") & Chr(9)
                s = s & ds!order_qty & Chr(9)
                s = s & ds!ship_plt_qty & Chr(9)
                s = s & ds!ship_status & Chr(9) & psz
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Form1.Check2.Value = 1 Then
        s = "select description,source,units,count(*) from paltasks"
        s = s & " where area = 'FORKLIFT' and status = 'PEND'"
        s = s & " and target = 'STAGING'"
        s = s & " and product >= '" & osku & "'"
        s = s & " and product < '" & osku & "ZZZ'"
        s = s & " group by description,source,units"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = Left(ds!Description, 6) & Chr(9)
                s = s & ds!Source & Chr(9)
                s = s & Format(Now, "m-d-yyyy") & Chr(9)
                s = s & ds(3) & Chr(9)
                s = s & "0" & Chr(9)
                s = s & "PEND" & Chr(9) & ds(2)
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Form1.plantno = "52" Then
        s = "select description,source,units,count(*) from paltasks"
        s = s & " where area = 'DOCK' and status = 'PEND'"
        s = s & " and source not in ('STAGING','ALT')"
        s = s & " and description > ' '"
        s = s & " and product >= '" & osku & "'"
        s = s & " and product < '" & osku & "ZZZ'"
        s = s & " and lotnum < '0'"
        s = s & " group by description,source,units"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = Left(ds!Description, 6) & Chr(9)
                s = s & ds!Source & Chr(9)
                s = s & Format(Now, "m-d-yyyy") & Chr(9)
                s = s & ds(3) & Chr(9)
                s = s & "0" & Chr(9)
                s = s & "PEND" & Chr(9) & ds(2)
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Form1.plantno = "50" Then
        s = "select description,source,units,count(*) from paltasks"
        s = s & " where area = 'DOCK' and status = 'PEND'"
        s = s & " and source in ('SR5', 'SR6')"             'jv082813
        s = s & " and description > ' '"
        s = s & " and product >= '" & osku & "'"
        s = s & " and product < '" & osku & "ZZZ'"
        s = s & " and lotnum < '0'"
        s = s & " group by description,source,units"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = Left(ds!Description, 6) & Chr(9)
                s = s & ds!Source & Chr(9)
                s = s & Format(Now, "m-d-yyyy") & Chr(9)
                s = s & ds(3) & Chr(9)
                s = s & "0" & Chr(9)
                s = s & "PEND" & Chr(9) & ds(2)
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    Grid1.FormatString = "^Order|^Whs|^Date|^Ordered|^Shipped|^Status|^Size"
    Grid1.ColWidth(0) = 800: Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 1100: Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800: Grid1.ColWidth(5) = 700
    Grid1.ColWidth(6) = 600
    Grid1.Redraw = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Form5.Caption = Form5.Caption & " " & Form1.plantdesc
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
            Form5.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form5.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form5.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form5.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    If Form5.Height > 3540 Then Grid1.Height = Form5.Height - 885
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    If Form5.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form5.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form5.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form5.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form5.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

Private Sub osku_Change()
    Call refresh_grid
End Sub


