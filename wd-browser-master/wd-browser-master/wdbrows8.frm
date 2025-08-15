VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13095
   LinkTopic       =   "Form8"
   ScaleHeight     =   8355
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   7215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5530
      _Version        =   327680
      Rows            =   5
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   65280
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Over Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30 Day Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Below 2 Week Level"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label brcode 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   735
   End
   Begin VB.Menu qmenu 
      Caption         =   "Lists"
      Begin VB.Menu qall 
         Caption         =   "All Products"
         Checked         =   -1  'True
      End
      Begin VB.Menu qW 
         Caption         =   "Below 2 Week Level"
      End
      Begin VB.Menu qY 
         Caption         =   "2 Week Supply"
      End
      Begin VB.Menu qb 
         Caption         =   "30 Day Supply"
      End
      Begin VB.Menu qG 
         Caption         =   "Overstocked"
      End
      Begin VB.Menu qpromo 
         Caption         =   "Promotions"
      End
      Begin VB.Menu qdisc 
         Caption         =   "Discontinued Products"
      End
      Begin VB.Menu qnosale 
         Caption         =   "No Recent Sales"
      End
      Begin VB.Menu qoo 
         Caption         =   "On Order"
      End
      Begin VB.Menu qpro 
         Caption         =   "Recommended Order"
      End
      Begin VB.Menu qpart 
         Caption         =   "Partial Pallet Order"
      End
      Begin VB.Menu dazestock 
         Caption         =   "Days In Stock"
      End
      Begin VB.Menu qsales 
         Caption         =   "Unit Sales"
         Begin VB.Menu qunit 
            Caption         =   "Unit Types"
            Visible         =   0   'False
         End
         Begin VB.Menu uhg 
            Caption         =   "1/2 Gallons"
         End
         Begin VB.Menu u48 
            Caption         =   "48 oz"
         End
         Begin VB.Menu upint 
            Caption         =   "Pints"
         End
         Begin VB.Menu uqt 
            Caption         =   "Quarts"
         End
         Begin VB.Menu u3gal 
            Caption         =   "3 Gallons"
         End
         Begin VB.Menu utray 
            Caption         =   "Trays"
         End
         Begin VB.Menu u6p 
            Caption         =   "6 Pack"
         End
         Begin VB.Menu u12p 
            Caption         =   "12 Pack"
         End
         Begin VB.Menu u24p 
            Caption         =   "24 Pack"
         End
         Begin VB.Menu ucup 
            Caption         =   "Cups"
         End
         Begin VB.Menu utake 
            Caption         =   "Take Home Snacks"
         End
         Begin VB.Menu ubulk 
            Caption         =   "Bulk Snacks"
         End
      End
   End
   Begin VB.Menu prtgrid 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub days_in_stock()
    Dim i As Integer, cfile As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String
    DoEvents
    cfile = Form1.webdir & "\brana\branches.csv"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
            If Val(brcode) = Val(f0) Then
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 0) = f1 Then
                        Grid1.TextMatrix(i, 2) = f3
                        Exit For
                    End If
                Next i
                DoEvents
            End If
        Loop
        Close #1
    End If
End Sub

Private Sub refresh_grid()
    Dim psku As String, pdesc As String, ppoh As String
    Dim puoh As String, poo As String, psales As String
    Dim udiff As String, pdiff As String, plnt As String
    Dim plit As String, pcc As String
    Dim s As String, pro As String, rpcode As String
    Dim ts As String
    Dim tc As Integer
    tc = Check1.Value
    Check1.Value = 0
    Screen.MousePointer = 11
    'ts = Form1.webdir
    'Form1.webdir = "c:\brana"
    Grid1.Visible = False
    Grid1.Clear: Grid1.Rows = 1
    Grid1.Cols = 13
    rpcode = brcode
    If brcode = "34" Then
        s = "Do you wish to view Sylacauga's inventory for the satellite routes?"
        If MsgBox(s, vbYesNo + vbQuestion, "Satellite Route Inventory...") = vbYes Then rpcode = "52"
    End If
    If Len(Dir(Form1.webdir & "\stock\gsales.R" & rpcode)) > 0 Then
        s = "The browser has determined that some of the routes associated with "
        s = s & "branch " & brcode & " do not actually load out at the branch itself."
        s = s & "  Click 'Yes' to view all route sales.  Click 'No' to view only the"
        s = s & " route loads from branch " & brcode & "'s cold storage warehouse."
        If MsgBox(s, vbYesNo, "Satellite Routes Detected.....") = vbYes Then rpcode = "R" & brcode
    End If
    If Len(Dir(Form1.webdir & "\stock\gsales." & rpcode)) > 0 Then
        s = "Branch " & rpcode & " Sales vs. Inventory"
        s = s & "  Last update: "
        s = s & Format(FileDateTime(Form1.webdir & "\stock\gsales." & rpcode), "m-d-yyyy h:mm am/pm")
        If rpcode <> brcode Then s = s & " All Routes"
        Form8.Caption = s
        Open Form1.webdir & "\stock\gsales." & rpcode For Input As #1
        Do Until EOF(1)
            Input #1, psku, pdesc, ppoh, puoh, poo, psales, udiff, pdiff, plnt, pro, plit, pcc
            If plnt = "50" Then plnt = "TX"
            If plnt = "51" Then plnt = "OK"
            If plnt = "52" Then plnt = "AL"
            s = psku & Chr(9) & pdesc & Chr(9) & Chr(9)
            s = s & Format(ppoh, "#") & Chr(9)
            s = s & Format(puoh, "#") & Chr(9)
            s = s & Format(poo, "#") & Chr(9)
            s = s & Format(psales, "#") & Chr(9)
            s = s & Format(udiff, "#") & Chr(9)
            s = s & Format(pdiff, "#") & Chr(9)
            s = s & plnt & Chr(9)
            s = s & pro & Chr(9)
            s = s & pcc & Chr(9)
            s = s & plit
            If qall.Checked Then Grid1.AddItem s
            If qW.Checked And pcc = "W" Then Grid1.AddItem s
            If qY.Checked And pcc = "Y" Then Grid1.AddItem s
            If qb.Checked And pcc = "B" Then Grid1.AddItem s
            If qG.Checked And pcc = "G" Then Grid1.AddItem s
            If qpromo.Checked And Left(plit, 5) = "Promo" Then Grid1.AddItem s
            If qdisc.Checked And Left(plit, 4) = "Disc" Then Grid1.AddItem s
            If qpro.Checked And Val(pdiff) < 0 Then Grid1.AddItem s
            If qpart.Checked And Val(udiff) < 0 And (Left(pdesc, 1) = "3" Or Left(pdesc, 2) = "TR") Then Grid1.AddItem s
            If qnosale.Checked And Val(psales) <= 0 And Val(puoh) > 0 Then Grid1.AddItem s
            If qoo.Checked And Val(poo) > 0 Then Grid1.AddItem s
            If qunit.Checked And Val(psales) > 0 Then
                If uhg.Checked And Left(pdesc, 3) = "1/2" Then Grid1.AddItem s
                If upint.Checked And Left(pdesc, 1) = "P" Then Grid1.AddItem s
                If ucup.Checked And Left(pdesc, 3) = "CUP" Then Grid1.AddItem s
                If ubulk.Checked And Left(pdesc, 3) = "BUL" Then Grid1.AddItem s
                If ubulk.Checked And Left(pdesc, 1) = "D" Then Grid1.AddItem s
                If u3gal.Checked And Left(pdesc, 1) = "3" Then Grid1.AddItem s
                If uqt.Checked And Left(pdesc, 1) = "Q" Then Grid1.AddItem s
                If utray.Checked And Left(pdesc, 2) = "TR" Then Grid1.AddItem s
                If u6p.Checked And Left(pdesc, 1) = "6" Then Grid1.AddItem s
                If u12p.Checked And Left(pdesc, 2) = "12" Then Grid1.AddItem s
                If u24p.Checked And Left(pdesc, 2) = "24" Then Grid1.AddItem s
                If u48.Checked And Left(pdesc, 2) = "48" Then Grid1.AddItem s
            End If
        Loop
    End If
    Close #1
    Grid1.FormatString = "^SKU|<Description|^Days In Stock|^Pallets OnHand|^Units OnHand|^Units OnOrder|^Sales Last 30 Days|^Units Diff|^Pallet Diff|^Plant|^ReOrder Pal Qty"
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 3800
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 900
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 900
    Grid1.ColWidth(11) = 1
    Grid1.ColWidth(12) = 1
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i: Grid1.Col = 0: Grid1.ColSel = 12
        If Grid1.TextMatrix(i, 11) = "W" Then Grid1.CellBackColor = wcolor.BackColor
        If Grid1.TextMatrix(i, 11) = "B" Then Grid1.CellBackColor = bcolor.BackColor
        If Grid1.TextMatrix(i, 11) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        If Grid1.TextMatrix(i, 11) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
        'Grid1.FillStyle = flexFillRepeat
        If Grid1.TextMatrix(i, 12) > "   " Then
            Grid1.Col = 0: Grid1.ColSel = 1
            'Grid1.CellFontUnderline = True
            Grid1.CellFontBold = True
            'Grid1.FillStyle = flexFillRepeat
        End If
    Next i
    'Grid1.Row = 1: Grid1.RowSel = 1: Grid1.Col = 0: Grid1.ColSel = 10
    'If Grid1.TextMatrix(1, 10) = "W" Then Grid1.CellBackColor = wcolor.BackColor
    'If Grid1.TextMatrix(1, 10) = "B" Then Grid1.CellBackColor = bcolor.BackColor
    'If Grid1.TextMatrix(1, 10) = "G" Then Grid1.CellBackColor = gcolor.BackColor
    'If Grid1.TextMatrix(1, 10) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
    'Grid1.FillStyle = flexFillRepeat
    
    If Grid1.Rows > 1 Then
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
        Grid1.Row = 1: Grid1.Col = 5
    End If
    If qpro.Checked Then
        Grid1.Col = 8: Grid1.ColSel = 8
        Grid1.RowSel = Grid1.Row
        Grid1.Sort = 3
    End If
    If qpart.Checked Then
        Grid1.Col = 7: Grid1.ColSel = 7
        Grid1.RowSel = Grid1.Row
        Grid1.Sort = 3
    End If
    If qunit.Checked Then
        Grid1.Col = 6: Grid1.ColSel = 6
        Grid1.RowSel = Grid1.Row
        Grid1.Sort = 4
    End If
    Grid1_RowColChange
    Grid1.Visible = True
    'Form1.webdir = ts
    Check1.Value = tc
    Call days_in_stock
    Screen.MousePointer = 0
End Sub

Private Sub brcode_Change()
    Call refresh_grid
End Sub

Private Sub dazestock_Click()
    Dim s As String, md As Boolean
    s = InputBox("Days in stock >=", "Days in stock...", "120")
    If Len(s) = 0 Then Exit Sub
    s = Val(s)
    qall_Click
    DoEvents
    For i = Grid1.Rows - 1 To 1 Step -1
        md = False
        If Val(Grid1.TextMatrix(i, 4)) < 1 Then
            md = True
        Else
            If Val(Grid1.TextMatrix(i, 2)) <= Val(s) Then
                md = True
            End If
        End If
        If md = True Then
            If Grid1.Rows > 2 Then
                Grid1.RemoveItem i
            Else
                Grid1.Rows = 1
            End If
        End If
    Next i
            
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form8.Width - 80
    If Form8.Height > 2000 Then Grid1.Height = Form8.Height - 930
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.saleinv.Checked = False
End Sub

Private Sub Grid1_DblClick()
    Dim rc As String
    If (Form1.wdbranch = "R1" And Me.u3gal.Checked = True) Or Form1.wdbranch = "SU" Then
        'rc = InputBox("Region Number (1-6):", "Specify Region..", "1")
        'If Len(rc) = 0 Then Exit Sub
        'If Val(rc) < 1 Or Val(rc) > 6 Then Exit Sub
        'Screen.MousePointer = 11
        'brwzbrana2.calledby = "Region " & rc
        brwzbrana2.calledby = "Region - All"
        brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
        Screen.MousePointer = 0
        brwzbrana2.Show
        Check1 = 1
        Exit Sub
    End If
    
    If Form1.wdbranch < "D1" Then Exit Sub
    If Form1.wdbranch > "D9" Then Exit Sub
    Screen.MousePointer = 11
    brwzbrana2.calledby = "Region " & Right(Form1.wdbranch, 1)
    brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    Screen.MousePointer = 0
    brwzbrana2.Show
    Check1 = 1
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer
    i = Grid1.Row
    If Check1 = 1 And Left(brwzbrana2.Caption, Len(Grid1.TextMatrix(i, 1))) <> Grid1.TextMatrix(i, 1) Then
        If Form1.wdbranch = "R1" Or Form1.wdbranch = "SU" Then
            brwzbrana2.calledby = "Region - All"
        Else
            brwzbrana2.calledby = "Region " & Right(Form1.wdbranch, 1)
        End If
        brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    End If
    If Len(Grid1.TextMatrix(Grid1.Row, 12)) > 0 And Grid1.TextMatrix(Grid1.Row, 12) > "   " Then
        Text1 = Grid1.TextMatrix(Grid1.Row, 12)
        Text1.Visible = True
    Else
        Text1 = ""
        Text1.Visible = False
    End If
End Sub

Private Sub Prtgrid_Click()
    Dim pl As String, i As Integer, lc As Integer
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    rh = "Branch " & brcode
    Screen.MousePointer = 11
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Screen.MousePointer = 0
    
    'lc = 4
    'Printer.Font = "Courier New"
    'Printer.FontSize = 8
    'Printer.Print Form8.Caption
    'Printer.Print " "
    ''Printer.Print "---------1---------2---------3---------4---------5---------6---------7---------8---------9---------A---------B"
    'Printer.Print "                                     Pallets     Units     Units    Sales      Unit     Pallet              Reorder"
    'Printer.Print " SKU                                  OnHand    OnHand    OnOrder  Last 30     Diff      Diff      Plant   Pallet Qty"
    'For i = 1 To Grid1.Rows - 1
    '    If lc > 76 Then
    '        Printer.NewPage
    '        lc = 4
    '        Printer.Print Form8.Caption
    '        Printer.Print " "
    '        Printer.Print "                                     Pallets     Units     Units    Sales      Unit     Pallet              Reorder"
    '        Printer.Print " SKU                                  OnHand    OnHand    OnOrder  Last 30     Diff      Diff      Plant   Pallet Qty"
    '    End If
    '    Grid1.TextMatrix(i, 1) = Left(Grid1.TextMatrix(i, 1), 30)
    '    pl = Space(4 - Len(Grid1.TextMatrix(i, 0))) & Grid1.TextMatrix(i, 0) & " "
    '    pl = pl & Grid1.TextMatrix(i, 1) & Space(30 - Len(Grid1.TextMatrix(i, 1))) & " "
    '    pl = pl & Space(7 - Len(Grid1.TextMatrix(i, 2))) & Grid1.TextMatrix(i, 2) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 3))) & Grid1.TextMatrix(i, 3) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 4))) & Grid1.TextMatrix(i, 4) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 5))) & Grid1.TextMatrix(i, 5) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 6))) & Grid1.TextMatrix(i, 6) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 7))) & Grid1.TextMatrix(i, 7) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 8))) & Grid1.TextMatrix(i, 8) & " "
    '    pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 9))) & Grid1.TextMatrix(i, 9)
    '    Printer.Print pl
    '    lc = lc + 1
    '    If Grid1.TextMatrix(i, 11) > "    " Then
    '        Printer.Print "     * " & Grid1.TextMatrix(i, 11)
    '        lc = lc + 1
    '    End If
    'Next i
    'Printer.EndDoc
    'Screen.MousePointer = 0
End Sub

Private Sub qall_Click()
    qall.Checked = True: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qB_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = True: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qdisc_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = True: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qG_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = True: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qnosale_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = True: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qoo_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = True
    refresh_grid
End Sub

Private Sub qpart_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = True
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qpro_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = True: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qpromo_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = True
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qunit_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = True: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qW_Click()
    qall.Checked = False: qW.Checked = True: qY.Checked = False
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub qY_Click()
    qall.Checked = False: qW.Checked = False: qY.Checked = True
    qb.Checked = False: qG.Checked = False: qpromo.Checked = False
    qdisc.Checked = False: qpro.Checked = False: qpart.Checked = False
    qunit.Checked = False: qnosale.Checked = False: qoo.Checked = False
    refresh_grid
End Sub

Private Sub u12p_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = True
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub u24p_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = True: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub u3gal_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = True: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub u48_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = True
    qunit_Click
End Sub

Private Sub u6p_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = True: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub ubulk_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = True
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub ucup_Click()
    uhg.Checked = False: ucup.Checked = True: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub uhg_Click()
    uhg.Checked = True: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub upint_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = True: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub uqt_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = True
    utray.Checked = False: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub

Private Sub utake_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = False: u6p.Checked = True: u12p.Checked = True
    u24p.Checked = True: utake.Checked = True: u48.Checked = False
    qunit_Click
End Sub

Private Sub utray_Click()
    uhg.Checked = False: ucup.Checked = False: ubulk.Checked = False
    upint.Checked = False: u3gal.Checked = False: uqt.Checked = False
    utray.Checked = True: u6p.Checked = False: u12p.Checked = False
    u24p.Checked = False: utake.Checked = False: u48.Checked = False
    qunit_Click
End Sub
