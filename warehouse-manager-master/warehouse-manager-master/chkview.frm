VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   6135
   ClientLeft      =   855
   ClientTop       =   750
   ClientWidth     =   7365
   LinkTopic       =   "Form10"
   ScaleHeight     =   6135
   ScaleWidth      =   7365
   Begin VB.ListBox reclist 
      Height          =   1035
      Left            =   5160
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   1000
      Left            =   0
      Max             =   12240
      SmallChange     =   1000
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   120
      Min             =   1
      TabIndex        =   7
      Top             =   0
      Value           =   1
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      LargeChange     =   1000
      Left            =   6840
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7095
   End
   Begin MSFlexGridLib.MSFlexGrid rkgrid 
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   327680
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label repdate 
      Caption         =   "repdate"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label reptrig 
      Caption         =   "reptrig"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label reptype 
      Caption         =   "reptype"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub view_fork_checkoff(pd As Control)
    Dim ds As ADODB.Recordset, sqlx As String
    Dim ss As ADODB.Recordset
    Dim rs As ADODB.Recordset, boxy As Long, testc As Integer
    Dim pdate As String, pl As String, pno As Integer
    Dim vl As Long, hl As Long
    Dim nsrc As Integer, nwhs As Integer, npal As Integer, ndesc As String
    Dim rsrc As String, rstr As String
    Dim rollerbed As Boolean
    Dim p As Integer
    If TypeOf pd Is PictureBox Then
        reclist.Clear
        pdate = Format(Now, "m-d-yyyy")
        pdate = InputBox("Receipt Date:", "Receipt Date...", pdate)
        repdate = Format(pdate, "m-d-yyyy")
    Else
        pdate = Format(repdate, "m-d-yyyy")
    End If
    If reptype = "RB" Then
        rollerbed = True
    Else
        rollerbed = False
    End If
    Screen.MousePointer = 11
    If TypeOf pd Is PictureBox Then
        pd.DrawWidth = 1
        pd.Cls
    Else
        pd.DrawWidth = 3
    End If
    pf = 14
    'pf = InputBox("Large Font Size:", "Font...", "14")
    If Val(pf) = 0 Then pf = 14
    If Val(pf) > 20 Then pf = 14
    pd.CurrentX = 0: pd.CurrentY = 0
    pd.FontName = "Courier New"
    pd.FontSize = pf
    pd.FontBold = True
    If rollerbed Then
        pd.Print "Rollerbed ";
    Else
        pd.Print "SR-4 ";
    End If
    pd.Print "Forklift Check-off Sheet  " & Format(pdate, "m-d-yyyy")
    pd.FontBold = False
    pd.FontSize = pf - 4
    sqlx = "select id,sku,proddate,units,lot_num,sr4"
    sqlx = sqlx & " from prodrcv"
    sqlx = sqlx & " where recdate1 = '" & pdate & "'"
    sqlx = sqlx & " or recdate2 = '" & pdate & "'"
    sqlx = sqlx & " or recdate3 = '" & pdate & "'"
    sqlx = sqlx & " order by proddate,sku"
    Set ds = Wdb.Execute(sqlx)
    p = 1
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If pd.CurrentY >= 12240 Then
                If TypeOf pd Is Printer Then
                    pd.NewPage
                Else
                    rstr = localAppDataPath & "\pic" & Format(p, "00000") & ".bmp"
                    SavePicture pd.Image, rstr
                    p = p + 1
                    pd.Cls
                End If
                pd.FontName = "Courier New"
                pd.FontSize = pf
                pd.FontBold = True
                If rollerbed Then
                    pd.Print "Rollerbed ";
                Else
                    pd.Print "SR-4 ";
                End If
                pd.Print "Forklift Check-off Sheet  " & Format(pdate, "m-d-yyyy")
            End If
            sqlx = "select * from skumast where sku = '" & ds!sku & "'"
            Set ss = Sdb.Execute(sqlx)
            If ss.BOF = True Then
                nsrc = 1: nwhs = 3: npal = 1: ndesc = "Invalid SKU"
            Else
                nsrc = ss!source
                nwhs = ss!whs_num
                npal = ss!pallet
                ndesc = ss!fgunit & " " & ss!fgdesc
            End If
            ss.Close
            If npal = 0 Then npal = 1
            pno = 0
            If rollerbed Then
                If nsrc = 2 And ds!sr4 = 0 Then pno = Int(ds!units / npal) + 1
            Else
                If ds!sr4 > 0 Then
                    If nwhs = 4 Then
                        pno = Int(ds!units / npal) + 1
                    Else
                        pno = ds!sr4
                    End If
                End If
            End If
            If pno > 0 And reptype = "User" Then
                If TypeOf pd Is PictureBox Then
                    If MsgBox("Print " & ndesc & "?", vbYesNo + vbQuestion, "product date " & ds!proddate & " Lot " & ds!lot_num) = vbNo Then
                        pno = 0
                    Else
                        reclist.AddItem ds!id
                    End If
                Else
                    rstr = ""
                    For i = 0 To reclist.ListCount - 1
                        If Val(reclist.List(i)) = ds!id Then
                            rstr = reclist.List(i)
                            Exit For
                        End If
                    Next i
                    If Len(rstr) = 0 Then pno = 0
                End If
            End If
            If pno > 0 Then
                pd.FontSize = pf - 4
                pd.FontBold = False
                pd.Print String(97, "-")
                pd.FontBold = True
                pd.Print ds!sku & " " & ndesc & "  " & Format(ds!proddate, "m-d-yyyy");
                pd.FontBold = False
                pd.PSet (6480 * (Val(pf) / 14), pd.CurrentY)
                pd.Print "Partial Wraps: |________|________|________|"
                hl = pd.TextWidth(" Tag #   01    02    03    04    05    06    07    08    09    10    11    12    13    14    15  ")
                pd.Line (0, pd.CurrentY)-(hl, pd.CurrentY)
                boxy = pd.CurrentY
                pd.Print ""
                pl = " Tag # "
                For i = 1 To pno
                    pl = pl & "  " & Format(i, "00") & "  "
                    If i Mod 15 = 0 Then
                        pd.Print pl
                        pd.Line (0, pd.CurrentY)-(hl, pd.CurrentY)
                        pd.Print " "
                        pl = "       "
                    End If
                Next i
                pd.Print pl
                pd.Line (0, pd.CurrentY)-(hl, pd.CurrentY)
                pd.Line (1, boxy)-(1, pd.CurrentY)
                vl = pd.TextWidth(" Tag # ")
                pd.Line (vl, boxy)-(vl, pd.CurrentY)
                For i = 1 To 15
                    vl = vl + pd.TextWidth("  01  ")
                    pd.Line (vl, boxy)-(vl, pd.CurrentY)
                Next i
                pd.Line (hl, boxy)-(hl, pd.CurrentY)
                pd.PSet (0, pd.CurrentY)
                pl = " "
                pd.Print pl
                hl = pd.TextWidth(" A Rrrrrrrrr     00 00  A RRRRRRRR     00 00  A RRRRRRRR     00 00   A RRRRRRRR     00 00   ")
                pd.PSet (0, pd.CurrentY)
                boxy = pd.CurrentY
                pd.Print " Rack       Beg   End   Rack       Beg   End   Rack       Beg   End   Rack       Beg   End "
                picy = pd.CurrentY
                pl = " "
                sqlx = "Select * from racks where resv_sku = '" & ds!sku & "'"
                sqlx = sqlx & " order by qty desc"
                Set rs = Wdb.Execute(sqlx)
                If rs.BOF = False Then
                    rs.MoveFirst
                    i = 1
                    Do Until rs.EOF
                        pl = pl & rs!aisle & " "
                        pl = pl & rs!rack & Space(8 - Len(rs!rack)) & " "
                        pl = pl & "            "
                        If i = 4 Then
                            pd.FontBold = True
                            pd.Print pl
                            pd.FontBold = False
                            pd.Line (0, pd.CurrentY)-(hl, pd.CurrentY)
                            pd.PSet (0, pd.CurrentY)
                            i = 1: pl = " "
                        End If
                        i = i + 1
                        rs.MoveNext
                    Loop
                End If
                rs.Close
                pd.FontBold = True
                pd.Print pl
                pd.FontBold = False
                pd.Line (0, pd.CurrentY)-(hl, pd.CurrentY)
                vl = pd.TextWidth(" A RRRRRRRR")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, boxy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" A RRRRRRRR")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, boxy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" A RRRRRRRR")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, boxy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" A RRRRRRRR")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, picy)-(vl, pd.CurrentY)
                vl = vl + pd.TextWidth(" 00   ")
                pd.Line (vl, boxy)-(vl, pd.CurrentY)
                pd.Print " "
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    rkgrid.Clear: rkgrid.Rows = 1
    sqlx = "select aisle,rack,slot,resv_sku,resv_lot from racks"
    sqlx = sqlx & " where resv_sku > '   ' or resv_lot > '     '"
    sqlx = sqlx & " order by aisle,slot,rack"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        pd.FontSize = pf - 4
        pd.FontBold = False
        pd.Print "Reserved Racks...."
        ds.MoveFirst
        rsrc = ds!aisle
        Do Until ds.EOF
            rstr = " " & ds!aisle & " "
            rstr = rstr & ds!rack & Space(8 - Len(ds!rack)) & " "
            rstr = rstr & ds!resv_sku & Space(5 - Len(ds!resv_sku)) & " "
            rstr = rstr & ds!resv_lot & Space(5 - Len(ds!resv_lot)) & "|"
            rkgrid.AddItem rstr
            ds.MoveNext
        Loop
        k = rkgrid.Rows - 1
        If Int(k / 4) <> k / 4 Then
            For i = k Mod 4 To 3
                rkgrid.AddItem "add"
            Next i
        End If
        k = (rkgrid.Rows - 1) / 4
        For i = 1 To k
            rkgrid.TextMatrix(i, 1) = rkgrid.TextMatrix(i + k, 0)
            rkgrid.TextMatrix(i, 2) = rkgrid.TextMatrix(i + k * 2, 0)
            rkgrid.TextMatrix(i, 3) = rkgrid.TextMatrix(i + k * 3, 0)
        Next i
        For i = rkgrid.Rows - 1 To 1 Step -1
            If rkgrid.TextMatrix(i, 3) = "add" Then rkgrid.TextMatrix(i, 3) = " "
            If rkgrid.TextMatrix(i, 3) = "" Then rkgrid.RemoveItem i
        Next i
        If rkgrid.Rows > 1 Then
            For i = 1 To rkgrid.Rows - 1
                rstr = rkgrid.TextMatrix(i, 0)
                rstr = rstr & rkgrid.TextMatrix(i, 1)
                rstr = rstr & rkgrid.TextMatrix(i, 2)
                rstr = rstr & rkgrid.TextMatrix(i, 3)
                pd.Print rstr
            Next i
        End If
    End If
    ds.Close ': db.Close: sb.Close
    If TypeOf pd Is Printer Then pd.EndDoc
    If TypeOf pd Is PictureBox Then
        If p > 1 Then
            rstr = localAppDataPath & "\pic" & Format(p, "00000") & ".bmp"
            SavePicture pd.Image, rstr
            pd.Picture = LoadPicture(localAppDataPath & "\pic00001.bmp")
            HScroll1.Visible = True
        Else
            HScroll1.Visible = False
        End If
        Form10.Caption = "Page 1 of " & p
        HScroll1.Max = p
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    Call view_fork_checkoff(Printer)
End Sub

Private Sub Form_Load()
    Picture1.Width = 12240
    Picture1.Height = 15840
    Me.Caption = "Page 1"
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        VScroll1.Height = Me.Height - 400
        VScroll1.Max = 15840 - VScroll1.Height
        VScroll1.SmallChange = Int(VScroll1.Max / 8)
        HScroll2.Top = Me.Height - 650
    End If
    If Me.Width > 2000 Then
        Frame1.Width = Me.Width
        VScroll1.Left = Me.Width - 380
        HScroll2.Width = Me.Width - 380
        HScroll2.Max = 12240 - HScroll2.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rstr = Dir(localAppDataPath & "\pic*.bmp")
    Do While Len(rstr) > 0
        Kill localAppDataPath & "\" & rstr
        rstr = Dir
    Loop
End Sub

Private Sub HScroll1_Change()
    rstr = localAppDataPath & "\pic" & Format(HScroll1.Value, "00000") & ".bmp"
    Picture1.Picture = LoadPicture(rstr)
    Me.Caption = "Page " & HScroll1.Value & " of " & HScroll1.Max
End Sub

Private Sub HScroll2_Change()
    Picture1.Move 0 - HScroll2.Value
End Sub

Private Sub reptrig_Change()
    Call view_fork_checkoff(Picture1)
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, Frame1.Height - VScroll1.Value
End Sub
