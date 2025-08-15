VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form browzbill 
   Caption         =   "Print Blank Bill of Laden"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13200
   LinkTopic       =   "Form3"
   ScaleHeight     =   8250
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
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
      Left            =   6720
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   120
      Width           =   4215
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
      Left            =   960
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   4215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      _Version        =   327680
      ForeColor       =   4210688
      BackColorFixed  =   12648447
      BackColorSel    =   32768
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      Caption         =   "Destination:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Origin:"
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
      Width           =   735
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu prtbill 
         Caption         =   "&Print Blank Bill"
      End
      Begin VB.Menu xitform 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "browzbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub duplex_bill()
    'Dim db As ADODB.Connection, ds As ADODB.Recordset, sqlx As String, s As String
    Dim sd As String, tp As String, tno As String, i As Integer, k As Integer
    Dim oaddr1 As String, oaddr2 As String, ophone As String, ofax As String
    Dim eno As Long, edesc As String
    If Grid1.Row < 1 Then Exit Sub
    tc = InputBox("Please Enter 4-digit decal Trailer Code # or 'OC' for Outside Carrier", "Trailer Code #", "OC")
    If Len(tc) = 0 Then Exit Sub
    'On Error GoTo SQLError
    'If UCase(tc) <> "OC" Then
    '    Set db = CreateObject("ADODB.Connection")
    '    'db.Open "ODBC;DATABASE=WDTruck;uid=bbctruck500;pwd=brenham500;DSN=truckwo"
    '    db.Open Form1.schdb
    '    s = "select listreturn from valuelists where listname = 'trlcode'"
    '    s = s & " and listreturn = '" & tc & "'"
    '    Set ds = db.Execute(s)
    '    If ds.BOF = True Then
    '        MsgBox "Invalid Trailer Code Entered", vbOKOnly + vbExclamation, "Sorry Cannot Process.."
    '        ds.Close: db.Close
    '        Exit Sub
    '    End If
    '    ds.Close: db.Close
    'End If
    sd = InputBox("Date:", Combo1 & " --> " & Combo2 & " Ship Date...", Format(Now, "m-d-yyyy"))
    If Len(sd) = 0 Then Exit Sub
    tp = InputBox("Total Pallets:", "Total Pallets...", "34")
    If Len(tp) = 0 Then Exit Sub
    'tno = InputBox("Trailer #:", "Trailer #....", "#1")
    'If Len(tno) = 0 Then Exit Sub
    tno = " " '"BT"
    Screen.MousePointer = 11
    
    'Printer.Duplex = 3
    'Printer.Orientation = 1
    
    'oplant = Form1.plantno
    k = 0
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = Combo1 Then
            k = i
            Exit For
        End If
    Next i
    oaddr1 = Grid1.TextMatrix(k, 1)
    oaddr2 = Grid1.TextMatrix(k, 2) & ", " & Grid1.TextMatrix(k, 3) & " " & Grid1.TextMatrix(k, 4)
    ophone = Grid1.TextMatrix(k, 5)
    ofax = Grid1.TextMatrix(k, 6)
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = Combo2 Then
            Grid1.Row = i
        End If
    Next i
    
    'Printer.Height = 1440 * 11
    'Printer.Width = 1440 * 8.5
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print Tab(32); " " '"B i l l   O f   L a d i n g"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentX = 720: Printer.Print "Origination:";
    Printer.FontBold = False
    Printer.CurrentX = 1440 * 1.5: Printer.Print "Blue Bell Creameries L.P.";
    Printer.FontBold = True
    Printer.CurrentX = 1440 * 4.5: Printer.Print "Destination: ";
    Printer.FontBold = False
    Printer.CurrentX = 1440 * 5.5
    Printer.Print Grid1.TextMatrix(Grid1.Row, 0); " "; tno
    
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr1; '"1101 S. Blue Bell Road";
    Printer.CurrentX = 1440 * 5.5
    Printer.Print Grid1.TextMatrix(Grid1.Row, 1)
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr2; '"Brenham, Texas  77834-1807";
    Printer.CurrentX = 1440 * 5.5
    Printer.Print Grid1.TextMatrix(Grid1.Row, 2); ", "; Grid1.TextMatrix(Grid1.Row, 3); " "; Grid1.TextMatrix(Grid1.Row, 4)
    Printer.CurrentX = 1440 * 1.5: Printer.Print ophone; '"(979) 836-7977";
    Printer.CurrentX = 1440 * 5.5
    Printer.Print Grid1.TextMatrix(Grid1.Row, 5)
    Printer.CurrentX = 1440 * 1.5: Printer.Print "Fax: " & ofax; '"Fax: (979) 830-7398";
    Printer.CurrentX = 1440 * 5.5
    If Grid1.TextMatrix(Grid1.Row, 6) > " " Then
        Printer.Print "Fax: "; Grid1.TextMatrix(Grid1.Row, 6)
    Else
        Printer.Print " "
    End If
    Printer.Print String(130, "_")
    
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentY = 1440 * 9
    'For i = lc To 50 '54 '45 '50 '57
    '    printer.Print " "
    'Next i
    Printer.CurrentX = 720: Printer.Print "Ship Date:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print sd;
    Printer.CurrentX = 1440 * 3: Printer.Print "Trailer #:";
    Printer.CurrentX = 1440 * 4: Printer.Print tc;
    Printer.CurrentX = 1440 * 5: Printer.Print "Total Pallets:";
    Printer.CurrentX = 1440 * 6: Printer.Print tp
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Inspected By:";                               'jv082415
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Completed By:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Seal #:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Sealed By:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Driver:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Freight:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Special Instructions:";
    Printer.CurrentX = 1440 * 2: Printer.Print "____________________________________________________________________"
    Printer.NewPage
    Call prtpage2(Printer)
    Printer.EndDoc
    'Printer.Duplex = 1
    Screen.MousePointer = 0
    Exit Sub
'SQLError:
'    eno = Err.Number: edesc = Err.Description: Err.Clear
'    Call vb_elog(eno, edesc, Me.Name, "duplex_bill", Form1.UserId)
'    If eno = -2147467259 Then
'        Resume
'    Else
'        MsgBox edesc, vbOKOnly, Me.Name & " duplex_bill - Error Number: " & eno
'        End
'    End If
End Sub

Private Sub prtpage2(pd As Control)
    Dim dl As String, s As String, i As Long
    Dim xs As Long, xe As Long, st As Long
    xs = 1440 * 0.25
    xe = 1440 * 8
    dl = "_________________________"
    'pd.Height = 1440 * 11
    'pd.Width = 1440 * 8.5
    pd.FontName = "Arial"
    pd.FontSize = 10
    If TypeOf pd Is Printer Then
        pd.DrawWidth = 6
    Else
        pd.DrawWidth = 1
    End If
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    s = "DRIVER INFORMATION"
    pd.FontBold = True
    pd.CurrentX = 1440 * 4 - (pd.TextWidth(s) * 0.5)
    pd.Print s
    pd.FontBold = False
    pd.Print " ": pd.Print " "
    st = pd.CurrentY
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 2.5: pd.Print "Driver #1";
    pd.CurrentX = 1440 * 4.5: pd.Print "Driver #2";
    pd.CurrentX = 1440 * 6.5: pd.Print "Driver #3"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Driver Name"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Starting Location"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Date"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Destination"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Depart temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Mid trip temp."             'jv022717
    pd.Print " "                                                    'jv022717
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)                     'jv022717
    pd.Print " "                                                    'jv022717
    pd.CurrentX = 1440 * 0.5: pd.Print "Arrival temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Signature"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 2: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 4: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 6: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 8: pd.Line (xs, st)-(xs, pd.CurrentY)
    pd.Print " "
    pd.Print " ": pd.Print " "
    s = "FINAL DESTINATION INFORMATION"
    pd.FontBold = True
    pd.CurrentX = 1440 * 4 - (pd.TextWidth(s) * 0.5)
    pd.Print s
    pd.FontBold = False

    
    pd.Print " ": pd.Print " "
    pd.CurrentX = 720: pd.Print "Arrival Date:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Arrival temperature:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Seal #:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Verified by:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Time Arrived:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Time Departed:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "# Pallets returned:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "# Sleeves returned:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Returns:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Comments:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Corrections:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Received by:";
    pd.CurrentX = 1440 * 2: pd.Print dl
End Sub

Private Sub refresh_grid()
    Dim cfile As String, f0 As String, f1 As String
    Dim f2 As String, f3 As String, f4 As String, f5 As String, f6 As String
    Dim s As String
    Combo1.Clear: Combo2.Clear
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    cfile = Form1.webdir & "\locflist.csv"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6
            Grid1.AddItem s
        Loop
        Close #1
    End If
    Grid1.FormatString = "<Ship To|<Address|<City|^State|^Zip|^Phone|^Fax"
    Grid1.ColWidth(0) = 2500
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 1500
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    If Grid1.Rows > 1 Then
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 1
        Grid1.Sort = 5
    End If
    For i = 1 To Grid1.Rows - 1
        Combo1.AddItem Grid1.TextMatrix(i, 0)
        Combo2.AddItem Grid1.TextMatrix(i, 0)
    Next i
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    s = UCase(Right(Form1.Combo1, Len(Form1.Combo1) - 5))
    'If Form1.plantno = "50" Then s = "BRENHAM"
    'If Form1.plantno = "51" Then s = "TULSA"
    'If Form1.plantno = "52" Then s = "SYLACAUGA"
    For i = 0 To Combo1.ListCount - 0
        If UCase(Combo1.List(i)) = s Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub Combo1_Click()
    Dim i As Integer
    For i = 0 To Grid1.Rows - 1
        If Trim(Grid1.TextMatrix(i, 0)) = Trim(Combo1) Then
            Grid1.Row = i: Grid1.TopRow = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            Exit For
        End If
    Next i
End Sub

Private Sub Combo2_Click()
    Dim i As Integer
    For i = 0 To Grid1.Rows - 1
        If Trim(Grid1.TextMatrix(i, 0)) = Trim(Combo2) Then
            Grid1.Row = i: Grid1.TopRow = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Width = Form1.Width
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height

    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1280 '880
End Sub

Private Sub Grid1_Click()
    Dim i As Integer
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = Grid1.TextMatrix(Grid1.Row, 0) Then
            Combo2.ListIndex = i
        End If
    Next i
End Sub

Private Sub prtbill_Click()
    'Dim db As Database, ds As Recordset, tc As String
    'Dim i As Integer, k As Integer, tp As String, tno As String
    If Grid1.Row < 1 Then Exit Sub
    Call duplex_bill
End Sub

Private Sub xitform_Click()
    Unload Me
End Sub
