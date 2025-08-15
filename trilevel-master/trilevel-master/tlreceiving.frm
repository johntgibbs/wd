VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tlreceiving 
   Caption         =   "Receiving List"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8700
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   8454143
      BackColorSel    =   192
      FocusRect       =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "tlreceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String
    Dim i As Integer, c As Boolean
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    'db.Open "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    s = "select p.id,p.sku,i.uom_type,i.description,p.proddate,p.lot_num,p.sr1,p.sr2,p.sr3,p.sr4,p.sr5,sp_flag"
    s = s & " from prodrcv p, sku_config i"
    s = s & " where i.sku = p.sku"
    's = s & " and p.sp_flag = '0'"
    s = s & " order by p.sr4,p.lot_num, p.sku"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)
            s = s & ds(1) & Chr(9)
            s = s & StrConv(ds(2), vbProperCase) & " " & StrConv(ds(3), vbProperCase) & Chr(9)
            s = s & Format(ds(4), "M-dd-yyyy") & Chr(9)
            s = s & ds(5) & Chr(9)
            s = s & Format(DateAdd("yyyy", 2, ds(4)), "MMddyy") & " " & ds(11) & Chr(9)
            s = s & Format(ds(6), "#") & Chr(9)
            s = s & Format(ds(7), "#") & Chr(9)
            s = s & Format(ds(8), "#") & Chr(9)
            s = s & Format(ds(9), "#") & Chr(9)
            s = s & Format(ds(10), "#") & Chr(9)
            If ds(9) > 0 Then
                s = s & "1"
            Else
                s = s & "0"
            End If
            s = s & ds(5) & ds(1)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    If Grid1.Rows > 1 Then
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 11: Grid1.ColSel = 11
        Grid1.Sort = 5
        Grid1.FillStyle = flexFillRepeat
        c = True
        For i = 1 To Grid1.Rows - 1
            c = Not c
            If c Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^Id|^SKU|<Product|^Date|^Lot|^Label|^SR-1|^SR-2|^SR-3|^SR-4|^SR-5"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4500
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1
    Grid1.Redraw = True
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Grid1.Font = "Arial"
    Grid1.FontSize = 10
    Grid1.FontBold = True
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 150
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1200
End Sub

