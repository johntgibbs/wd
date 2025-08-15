VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimprtloads 
   Caption         =   "Daily Route Loads"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   10680
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   10680
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8520
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox ldate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      _Version        =   327680
      ForeColor       =   4210688
      BackColorFixed  =   12648447
   End
   Begin VB.Label plit 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7920
      TabIndex        =   13
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label blit 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "SKU:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Route:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Branch;"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Load Date:"
      BeginProperty Font 
         Name            =   "Arial"
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
      Width           =   975
   End
   Begin VB.Menu printmenu 
      Caption         =   "Print"
      Begin VB.Menu prtgrid 
         Caption         =   "Print Grid Listing"
      End
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Begin VB.Menu postcnt 
         Caption         =   "Post to Countsheets"
      End
   End
End
Attribute VB_Name = "bimprtloads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo1.AddItem Format(branchrec(i).branchno, "000")
            If i = 1 Then                                       'jv090216
                List1.AddItem "Brenham Sales"                   'jv090216
            Else                                                'jv090216
                If i = 47 Then                                  'jv090216
                    List1.AddItem "Tulsa Sales"                 'jv090216
                Else                                            'jv090216
                    If i = 52 Then                              'jv090216
                        List1.AddItem "Sylacauga Sales"         'jv090216
                    Else                                        'jv090216
                        List1.AddItem branchrec(i).branchname
                    End If                                      'jv090216
                End If                                          'jv090216
            End If                                              'jv090216
        End If                                                  'jv090216
    Next i                                                      'jv090216
    Combo2.Clear: List2.Clear
    Combo2.AddItem "ALL"
    List2.AddItem "All Products"
    For i = 0 To 9999
        If skurec(i).sku > "0" Then
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
        End If
    Next i
    Combo3.Clear
    Combo3.AddItem "ALL"
    For i = 0 To 99
        Combo3.AddItem Format(i, "00")
    Next i
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
End Sub

Sub refresh_grid()
    Dim ds As ADODB.Recordset, q As String, s As String
    Dim i As Integer, t As Long, j As Long
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
        
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 5: Grid1.Rows = 1
    Grid1.FixedCols = 1
    Grid1.Clear
    
    q = "select product_no,route_no,sum(tran_qty)"
    q = q & " from bolinf.inv_adj_input_detail"
    q = q & " where tran_type = '1'"
    q = q & " and tran_date = TO_DATE('" & Format(ldate, "dd-mmm-yy") & "')"
    q = q & " and branch_no = '" & Combo1 & "'"
    If Combo2 <> "ALL" Then q = q & " and product_no = '" & Combo2 & "'"
    If Combo3 <> "ALL" Then q = q & " and route_no = '" & Combo3 & "'"
    q = q & " group by product_no, route_no"
    q = q & " order by product_no, route_no"
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!route_no & Chr(9)
            s = s & ds!product_no & Chr(9)
            s = s & skurec(Val(ds!product_no)).unit & " "
            s = s & skurec(Val(ds!product_no)).desc & Chr(9)
            s = s & ds(2) & Chr(9)
            j = skurec(Val(ds!product_no)).wrapunits
            If j > 0 And ds(2) >= j Then
                s = s & Format(ds(2) / skurec(Val(ds!product_no)).wrapunits, "0")
            End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If (Combo2 <> "ALL" Or Combo3 <> "ALL") And Grid1.Rows > 1 Then
        t = 0: j = 0
        For i = 1 To Grid1.Rows - 1
            t = t + Val(Grid1.TextMatrix(i, 3))
            j = j + Val(Grid1.TextMatrix(i, 4))
        Next i
        s = " " & Chr(9) & " " & Chr(9) & "Totals" & Chr(9) & t & Chr(9) & j
        Grid1.AddItem s
    End If
    Grid1.FormatString = "^Route|^SKU|<Product|^Units|^Wraps"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 1
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    blit.Caption = List1
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    plit.Caption = List2
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    ldate = Format(DateAdd("d", -1, Now), "M-dd-yyyy")
    refresh_vlists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 5)
End Sub

Private Sub postcnt_Click()
    Dim i As Integer, s As String, cfile As String
    If Grid1.Rows < 2 Then Exit Sub
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\routes." & lbr
    'cfile = "u:\loadtest.txt"
    cfile = InputBox("Route Countsheet File:", "Export File..", cfile)
    If Len(cfile) = 0 Then Exit Sub
    If Len(Dir(cfile)) > 0 Then
        s = "Filename: " & cfile & " already exists.  Select 'Yes' to append or 'No' for new file."
        If MsgBox(s, vbYesNo + vbQuestion, "Append to existing countsheet file....") = vbYes Then
            Open cfile For Append As #1
        Else
            Open cfile For Output As #1
        End If
    End If
    Screen.MousePointer = 11
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 1)) > 0 Then
            Write #1, "RT" & Grid1.TextMatrix(i, 0);
            Write #1, Grid1.TextMatrix(i, 1);
            Write #1, Grid1.TextMatrix(i, 2);
            Write #1, ""; "";
            Write #1, Val(Grid1.TextMatrix(i, 3)) * -1;
            Write #1, Val(Grid1.TextMatrix(i, 3)) * -1;
            Write #1, Format(ldate, "m-d-yyyy")
        End If
    Next i
    Close #1
    Screen.MousePointer = 0
End Sub

Private Sub prtgrid_Click()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer
    rt = Me.Caption & " " & Combo1 & "-" & List1
    rh = "Load Date: " & ldate & " Route: " & Combo3 & " SKU: " & Combo2
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "Yellow": gndc(0) = Grid1.BackColorFixed
    'htdc(0) = "Pink": gndc(0) = Grid1.BackColorFixed
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Grid1.Redraw = False
        Call htmlcolorgrid(Me, "c:\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        Grid1.Redraw = True
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub
