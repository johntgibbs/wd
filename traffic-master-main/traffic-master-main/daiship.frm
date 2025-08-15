VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form daiship 
   Caption         =   "Daifuku Shipping Orders"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   6240
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post Order"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   9120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   7800
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   6165
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Trailers:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "daiship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub dai_ship_message()
    Dim s As String, i As Integer, cfile As String, xname As String, rkey As Long
    If Grid1.Rows < 2 Then Exit Sub
    xname = "OrderItemMessage"
    cfile = "c:\jvwork\dai" & xname & ".xml"
    Open cfile For Output As #1
    s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
    s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
    s = s & "<!DOCTYPE OrderItemMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
    s = s & "<OrderItemMessage>" & vbCrLf
    s = s & "<Order action=" & Chr(34) & "ADD" & Chr(34)
    s = s & " sOrderID=" & Chr(34) & Trim(Left(Combo1, 8)) & Chr(34)
    s = s & " iPriority=" & Chr(34) & "7" & Chr(34)
    s = s & " iOrderStatus= " & Chr(34) & "READY" & Chr(34) & ">" & vbCrLf
    Print #1, s
    
    s = "<OrderHeader>" & vbCrLf
    s = s & "<sDestinationStation>" & List2 & "</sDestinationStation>" & vbCrLf
    s = s & "<sDescription>" & List1 & "</sDescription>" & vbCrLf
    s = s & "<sOrderMessage/>" & vbCrLf
    s = s & "</OrderHeader>" & vbCrLf
    Print #1, s
    
    For i = 1 To Grid1.Rows - 1
        s = "<OrderLine sItem=" & Chr(34) & Trim(Left(Grid1.TextMatrix(i, 4), 4)) & Chr(34) & ">" & vbCrLf
        s = s & "<sRouteID/>" & vbCrLf
        s = s & "<fOrderQuantity>" & Val(Grid1.TextMatrix(i, 5)) & "</fOrderQuantity>" & vbCrLf
        s = s & "<sDescription>" & Right(Grid1.TextMatrix(i, 4), Len(Grid1.TextMatrix(i, 4)) - 4) & "</sDescription>" & vbCrLf
        s = s & "</OrderLine>" & vbCrLf
        Print #1, s
    Next i
    
    s = "</Order>" & vbCrLf
    s = s & "</OrderItemMessage>"
    Print #1, s
    Close #1
    DoEvents
    rkey = wd_seq("DAIRequests")
    Call write_oracle_request("OrderItemMessage", rkey)
    DoEvents
    WebBrowser1.Navigate cfile
    
    'If d.sStoreDestination = "2" Or d.sStoreDestination = "3" Or d.sStoreDestination = "5" Then
    '    Text1.Text = Dai_expected_receipt(d)
    '    Open "c:\jvwork\daiExpectedReceiptMessage.xml" For Output As #1
    '    Print #1, Text1.Text
    '    Close #1
    '    DoEvents
    '    Call write_oracle_request("ExpectedReceiptMessage", Val(d.sOrderID))
    '    WebBrowser1.Navigate2 "c:\jvwork\daiExpectedReceiptMessage.xml"
    'End If
    
End Sub

Sub refresh_groups()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Combo1.Clear: List1.Clear: List2.Clear: List3.Clear
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select product, source, target from paltasks"
    s = s & " where area = 'GROUP'"
    s = s & " and target > ' '"
    s = s & " and source in (select target from paltasks"
    s = s & " where area = 'DOCK' and status = 'PEND' and userid <= ' ' and source = 'SR3')"
    's = s & " and status = 'PEND' order by trandate"
    s = s & " order by trandate"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!product
            List1.AddItem ds!source
            List3.AddItem Left(ds!product, 6)
            If ds!target = "DOOR5" Or ds!target = "DOOR6" Then
                List2.AddItem "2204"
            Else
                If ds!target = "DOOR3" Or ds!target = "DOOR4" Then
                    List2.AddItem "2205"
                Else
                    List2.AddItem "2207"
                End If
            End If
            'List2.AddItem ds!target
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Sub refresh_grid1_sum()
    Dim db As ADODB.Connection, ds As Recordset, s As String, psku As String
    Dim ss As Recordset, q As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select area,description,source,target,product,count(*) from paltasks"
    s = s & " where area = 'DOCK'"
    's = s & " and description = '" & Trim(Left(Combo1, 8)) & "'"
    s = s & " and description = '" & Trim(List3) & "'"
    s = s & " and source = 'SR3'"
    's = s & " and target = '" & List1 & "'"
    s = s & " and status = 'PEND'"
    s = s & " and userid <= ' '"
    s = s & " group by area,description,source,target,product"
    s = s & " having count(*) > 0"
    s = s & " order by area,description,source,target,product"
    MsgBox s
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!area & Chr(9)
            s = s & ds!description & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds(5)
            psku = Left(ds!product, 3)
            q = "select sku from lane where whse_num = 5"
            q = q & " and sku = '" & psku & "'"
            Set ss = db.Execute(q)
            If ss.BOF = False Then
            'If psku = "368" Or psku = "441" Then
                Grid1.AddItem s
            End If
            ss.Close
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "<Area|<Group Code|<Source|<Target|<Product|^Qty"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1800
    Grid1.ColWidth(4) = 1800
    Grid1.ColWidth(5) = 1800
End Sub

Sub refresh_grid1()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 17
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select * from paltasks"
    s = s & " where area = 'DOCK'"
    s = s & " and source = 'SR3'"
    s = s & " and target = '" & List1 & "'"
    s = s & " and status = 'PEND'"
    s = s & " and userid <= ' '"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & ds!description & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lotnum2 & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!trandate & Chr(9)
            s = s & ds!reqid
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^Id|<Area|<Desc|<Source|<Target|<Product|<BarCode|^Qty|^Uom|^Lot|^Units|^Lot2|^Units|^Status|^User|^Time|<Reqid"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1800
    Grid1.ColWidth(4) = 1800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 1800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1600
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 800
    Grid1.ColWidth(14) = 800
    Grid1.ColWidth(15) = 800
    Grid1.ColWidth(16) = 800
End Sub

Private Sub Combo1_Click()
    List3.ListIndex = Combo1.ListIndex
    List2.ListIndex = Combo1.ListIndex
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Call dai_ship_message
End Sub

Private Sub Form_Load()
    refresh_groups
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    WebBrowser1.Width = Me.Width - 80
End Sub

Private Sub List1_Click()
    refresh_grid1_sum
End Sub
