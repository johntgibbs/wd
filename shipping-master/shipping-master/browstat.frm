VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form browstat 
   Caption         =   "W/D Browser Status - What's going on.....What's going on..."
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form3"
   ScaleHeight     =   5865
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   327680
      ForeColor       =   4210688
      BackColorFixed  =   12648447
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label ccolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "browstat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    sqlx = "select max(rct_date) from brhist"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Import Trailer Info" & Chr(9)
        sqlx = sqlx & Format(ds(0), "m-d-yyyy")
        Grid1.AddItem sqlx
    End If
    ds.Close
    
    'sqlx = "Trailer Status Reports" & Chr(9)
    'sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\trlstat.04"), "m-d-yyyy h:mm am/pm")
    'Grid1.AddItem sqlx
    sqlx = "Out-of-Stock Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock.htm"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "New Release Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\release.txt"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Available Files" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\avord.28"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Sales vs Inventory Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\gsales.28"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Pallet Orders Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\ro5028.txt"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Partial Pallet Orders Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\rpart28.txt"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Sylacauga Plant Totals" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\gsales.502"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Broken Arrow Plant Totals" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\gsales.501"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Brenham Plant Totals" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\gsales.500"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Snack Plant Totals" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\gsales.505"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Oracle Inventory Report" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\stock\goh.28"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    sqlx = "Browser Prep - BIMP" & Chr(9)
    sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\brana\branches.csv"), "m-d-yyyy h:mm am/pm")
    Grid1.AddItem sqlx
    
    sqlx = "Old Product Report Due - "
    sqlx = sqlx & Format(DateAdd("d", 30, FileDateTime("s:\wd\data\oldpprt.txt")), "m-d-yyyy")
    If Format(DateAdd("d", 30, FileDateTime("s:\wd\data\oldpprt.txt")), "yyyymmdd") <= Format(Now, "yyyymmdd") Then
        sqlx = sqlx & Chr(9) & Format(FileDateTime("s:\wd\data\oldpprt.txt"), "m-d-yyyy h:mm am/pm")
    End If
    Grid1.AddItem sqlx
    If Len(Dir(Form1.webdir & "\orderoff.txt")) > 0 Then
        Grid1.AddItem "Branch Orders - Off"
    Else
        Grid1.AddItem "Branch Orders - Enabled"
    End If
    Grid1.FormatString = "<Process|^Last Done"
    Grid1.ColWidth(0) = 5000
    Grid1.ColWidth(1) = 2000
    For i = 0 To Grid1.Rows - 1
        If IsDate(Grid1.TextMatrix(i, 1)) Then
            If Format(Grid1.TextMatrix(i, 1), "yyyymmdd") < Format(Now, "yyyymmdd") Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.FillStyle = flexFillSingle
                Grid1.CellBackColor = ccolor.BackColor
                DoEvents
                Grid1.FillStyle = flexFillRepeat
            End If
        End If
    Next i
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
Private Sub Form_Load()
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = browstat.Width - 110
    If browstat.Height > 2000 Then Grid1.Height = browstat.Height - 680
End Sub
