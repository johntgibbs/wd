VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzings 
   Caption         =   "Branch Ingredient Storage"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11385
   LinkTopic       =   "Form13"
   ScaleHeight     =   6195
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7646
      _Version        =   327680
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label brcode 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "brwzings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ofile As String, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Grid1.FormatString = "^Whs|<Location|<Item|<Lot|^OnHand Qty|^Pkgs|^Pallets|^Tran Dates|^Qty|^Pkgs|^Pallets|^Reason"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 3200
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 800
    ofile = Form1.webdir & "\stock\ingstore." & Me.brcode
    If Len(Dir(ofile)) = 0 Then Exit Sub
    Me.Caption = "Oracle Ingredient Tracking - Updated: " & Format(FileDateTime(ofile), "m-d-yyyy h:mm am/pm")
    Open ofile For Input As #1
    Do Until EOF(1)
        Line Input #1, s
        Grid1.AddItem s
    Loop
    Close #1
End Sub

Private Sub brcode_Change()
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 740
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = "Plant - " & Me.brcode
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

