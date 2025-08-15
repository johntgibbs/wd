VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form daimessage 
   Caption         =   "Daifuku Messages"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "daimessage.frx":0000
      Top             =   4800
      Width           =   8895
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      ExtentX         =   16960
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Daifuku Message Types:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "daimessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadDocument(docfile As String)
    Dim xdoc As MSXML2.DOMDocument60
    Set xdoc = New MSXML2.DOMDocument60
    xdoc.validateOnParse = False
    If xdoc.Load(docfile) Then
    ' The document loaded successfully.
    ' Now do something intersting.
        Text2 = xdoc.documentElement.nodeName & vbCrLf
        DisplayNode xdoc.childNodes, 0
    Else
        MsgBox " The document failed to load."
        ' See the previous listing for error information.
    End If
End Sub

Public Sub DisplayNode(ByRef Nodes As IXMLDOMNodeList, ByVal Indent As Integer)

    Dim xNode As IXMLDOMNode
    Dim xattr As IXMLDOMAttribute
    Indent = Indent + 2

    For Each xNode In Nodes
        If xNode.nodeType = NODE_TEXT Then
            Text2 = Text2 & Space$(Indent) & xNode.parentNode.nodeName & _
            ":" & xNode.nodeValue & vbCrLf '" type: " & xNode.nodeType & vbCrLf
        Else
            'Text2 = Text2 & xNode.parentNode.nodeName & " type: " & xNode.nodeType & vbCrLf
            'Text2 = Text2 & xNode.nodeName & " type: " & xNode.nodeType & vbCrLf
        End If
      
        If xNode.nodeType = 1 Then
            If xNode.Attributes.length > 0 Then
                For i = 0 To xNode.Attributes.length - 1
                    'Text2.Text = Text2.Text & "attr(" & i & "): " & xNode.Attributes(i).nodeName
                    Text2.Text = Text2.Text & xNode.nodeName & " " & xNode.Attributes(i).nodeName & _
                    "=" & xNode.Attributes(i).nodeValue & vbCrLf
                Next i
            'Else
            '    Text2 = Text2 & xNode.nodeName & vbCrLf
            End If
        End If
        
      If xNode.hasChildNodes Then
         DisplayNode xNode.childNodes, Indent
      End If
   Next xNode
End Sub

Private Sub save_xml_grid(gc As Integer, rootname As String, rowname As String)
    Dim xmldoc As MSXML2.DOMDocument60
    Dim ProcInstr As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim rowelement As IXMLDOMElement
    Dim cElement As IXMLDOMElement
    Dim dElement As IXMLDOMElement
    Dim gElement() As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim i As Integer, k As Integer, s As String
    ReDim gElement(gc)
    'Creating DOM Document object
    Set xmldoc = New MSXML2.DOMDocument60
    'this adds the processing instruction
    'the first line in an XML document

    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmldoc.appendChild ProcInstr
    'Create the root element
    Set rootElement = xmldoc.createElement(rootname)
    Set xmldoc.documentElement = rootElement
    'Creating comment node
    Set comElement = xmldoc.createComment("Daifuku " & Combo1 & " message.")
    'add the comment node after the root
    rootElement.appendChild comElement
    
    For i = 1 To Grid1.Rows - 1
        'Create the node student
        Set rowelement = xmldoc.createElement(rowname)
        'add the student node to the root
        rootElement.appendChild rowelement
        For k = 0 To gc - 1
            ''create a child element, 'name' for the student
            'Set cElement = xmldoc.createElement(Grid1.TextMatrix(0, 1))
            ''add a value for this node, 'Linda Jones'
            'cElement.nodeTypedValue = Grid1.TextMatrix(i, 1)
            ''cElement is child of aElement
            ''add the cElement as a child of aElement
            'aElement.appendChild cElement
            ''append another child to the 'student' element
            ''-------------
            'Set dElement = xmldoc.createElement(Grid1.TextMatrix(0, 2))
            'dElement.nodeTypedValue = Grid1.TextMatrix(i, 2)
            'aElement.appendChild dElement
            
            Set gElement(k) = xmldoc.createElement(Grid1.TextMatrix(0, k))
            gElement(k).nodeTypedValue = Grid1.TextMatrix(i, k)
            rowelement.appendChild gElement(k)
        Next k
        ''-------attributes---
        ''---create an attribute using the createAttribute() method
        ''---at the same time set its name
        'Set att = xmldoc.createAttribute("seq")
        ''---set the attributes 'text' property
        'att.Text = i 'Grid1.TextMatrix(i, 0)
        ''--for the aElement, which is 'student' set a named
        ''---attribute, it's name is att
        'aElement.Attributes.setNamedItem att
    Next i
    
    s = localAppDataPath & "\dai" & rootname & ".xml"
    xmldoc.Save s
    WebBrowser1.Navigate2 (s)
    DoEvents
    Text2 = ""
    Call LoadDocument(s)
End Sub

Public Sub dai_poll_messages()
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim xmname As String, seqid As Long, s As String, cfile As String
    cfile = "c:\jvwork\messages.txt"
    Open cfile For Append As #9
    seqid = 0
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.daisqldb
    sqlx = "SELECT iMessageSequence, sMessageIdentifier FROM WrxToHost ORDER BY iMessageSequence"
    Set ds = db.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            xmname = ds!smessageidentifier
            seqid = ds!imessagesequence
            Call read_dai_message(xmname, seqid)
            DoEvents
            s = "c:\jvwork\dai" & xmname & ".xml"
            WebBrowser1.Navigate2 (s)
            DoEvents
            Text2 = ""
            Call LoadDocument(s)
            'MsgBox "Sequence: " & seqid, vbOKOnly + vbInformation, xmname
            DoEvents
            Print #9, seqid
            Print #9, Text2.Text
            sqlx = "DELETE FROM WrxToHost WHERE iMessageSequence = " & seqid
            db.Execute sqlx
            If MsgBox("Sequence: " & seqid & "  Continue?", vbYesNo + vbQuestion, xmname) = vbNo Then Exit Do
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Close #9
    'If seqid > 0 Then
    '    Call read_dai_message(xmname, seqid)
    '    DoEvents
    '    s = "c:\jvwork\dai" & xmname & ".xml"
    '    WebBrowser1.Navigate2 (s)
    '    DoEvents
    '    Text2 = ""
    '    Call LoadDocument(s)
    'End If
End Sub

Sub dai_message_error()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "3" & Chr(9) & "1237" & Chr(9) & "Error adding record" & Chr(9)
    s = s & "OrderMessage::sOrderID missing in OrderHeader"
    Grid1.AddItem s
    s = "iErrorCode|"
    s = s & "iOriginalSequence|"
    s = s & "sHostErrorText|"
    s = s & "sMessage"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "ErrorMessage", "Error")
    'DoEvents
    'Call save_oracle_clob_message("ErrorMessage", 1022)
    'DoEvents
    'Call read_oracle_clob_message("ErrorMessage", 1022)
End Sub

Sub dai_message_inv_adj()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "06/05/2004 14:38:42" & Chr(9)
    s = s & "742" & Chr(9)
    s = s & "12277" & Chr(9)
    s = s & "-16" & Chr(9)
    s = s & "T30001" & Chr(9)
    s = s & "Cycle Count" & Chr(9)
    s = s & "RonG"
    Grid1.AddItem s
    s = "dTransactionTime|"
    s = s & "sItem|"
    s = s & "sLot|"
    s = s & "fAdjustQuantity|"
    s = s & "sLoadID|"
    s = s & "sReasonCode|"
    s = s & "sUserID"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "InventoryAdjustmentMessage", "InventoryAdjustment")
    DoEvents
    'Call save_oracle_clob_message("InventoryAdjustmentMessage", 1021)
    'DoEvents
    'Call read_oracle_clob_message("InventoryAdjustmentMessage", 1021)
End Sub

Sub dai_message_inv_status()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    s = "507" & Chr(9)
    s = s & "12305" & Chr(9)
    s = s & "60" & Chr(9)
    s = s & "T30001" & Chr(9)
    s = s & "INS" & Chr(9)
    s = s & "BrianJ"
    Grid1.AddItem s
    s = "sItem|"
    s = s & "sLot|"
    s = s & "fQuantity|"
    s = s & "sLoadID|"
    s = s & "sHoldReason|"
    s = s & "sUserID"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "InventoryStatusMessage", "InventoryStatus")
    DoEvents
    'Call save_oracle_clob_message("InventoryStatusMessage", 1020)
    'DoEvents
    'Call read_oracle_clob_message("InventoryStatusMessage", 1020)
End Sub

Sub dai_message_inv_upload()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    s = "SR5" & Chr(9)
    s = s & "924" & Chr(9)
    s = s & "12260" & Chr(9)
    s = s & "1120" & Chr(9)
    s = s & "INS" & Chr(9)
    Grid1.AddItem s
    s = "sWarehouse|"
    s = s & "sItem|"
    s = s & "sLot|"
    s = s & "fQuantity|"
    s = s & "sHoldReason"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "InventoryUploadMessage", "InventoryUpload")
    DoEvents
    'Call save_oracle_clob_message("InventoryUploadMessage", 1019)
    'DoEvents
    'Call read_oracle_clob_message("InventoryUploadMessage", 1019)
End Sub

Sub dai_message_store_comp()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    s = "06/05/2004 14:38:42" & Chr(9)
    s = s & "BH50216" & Chr(9)
    s = s & "102" & Chr(9)
    s = s & "12277" & Chr(9)
    s = s & "208" & Chr(9)
    s = s & "T30001" & Chr(9)
    s = s & "Dock" & Chr(9)
    s = s & "GaryH"
    Grid1.AddItem s
    s = "dTransactionTime|"
    s = s & "sOrderID|"
    s = s & "sItem|"
    s = s & "sLot|"
    s = s & "fReceivedQuantity|"
    s = s & "sLoadID|"
    s = s & "sStationName|"
    s = s & "sUserID"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "StoreCompleteMessage", "StoreComplete")
    DoEvents
    'Call save_oracle_clob_message("StoreCompleteMessage", 1018)
    'DoEvents
    'Call read_oracle_clob_message("StoreCompleteMessage", 1018)
End Sub

Sub dai_message_load_arrival()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "06/05/2004 14:38:42" & Chr(9)
    s = s & "BH50216" & Chr(9)
    s = s & "T30001" & Chr(9)
    s = s & "Dock"
    Grid1.AddItem s
    s = "dTransactionTime|"
    s = s & "sOrderID|"
    s = s & "sLoadID|"
    s = s & "sStationName"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "LoadArrivalMessage", "LoadArrival")
    DoEvents
    'Call save_oracle_clob_message("LoadArrivalMessage", 1017)
    'DoEvents
    'Call read_oracle_clob_message("LoadArrivalMessage", 1017)
End Sub

Sub dai_message_pick_comp()
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    s = "06/05/2004 14:38:42" & Chr(9)
    s = s & "BH50216" & Chr(9)
    s = s & "102" & Chr(9)
    s = s & "12277" & Chr(9)
    s = s & "208" & Chr(9)
    s = s & "T30001" & Chr(9)
    s = s & "Door3" & Chr(9)
    s = s & "BrentJ"
    Grid1.AddItem s
    s = "dTransactionTime|"
    s = s & "sOrderID|"
    s = s & "sItem|"
    s = s & "sLot|"
    s = s & "fPickQuantity|"
    s = s & "sLoadID|"
    s = s & "sStationName|"
    s = s & "sUserID"
    Grid1.FormatString = s
    Call save_xml_grid(Grid1.Cols, "PickCompleteMessage", "PickComplete")
    DoEvents
    'Call save_oracle_clob_message("PickCompleteMessage", 1016)
    'DoEvents
    'Call read_oracle_clob_message("PickCompleteMessage", 1016)
End Sub

Private Sub Combo1_Click()
    If Combo1 = "Error" Then dai_message_error
    If Combo1 = "Inventory Adjustment" Then dai_message_inv_adj
    If Combo1 = "Inventory Status" Then dai_message_inv_status
    If Combo1 = "Inventory Upload" Then dai_message_inv_upload
    If Combo1 = "Store Complete" Then dai_message_store_comp
    If Combo1 = "Load Arrival" Then dai_message_load_arrival
    If Combo1 = "Pick Complete" Then dai_message_pick_comp
    If Combo1 = "Poll Messages" Then dai_poll_messages
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Error"
    Combo1.AddItem "Inventory Adjustment"
    Combo1.AddItem "Inventory Status"
    Combo1.AddItem "Inventory Upload"
    Combo1.AddItem "Store Complete"
    Combo1.AddItem "Load Arrival"
    Combo1.AddItem "Pick Complete"
    Combo1.AddItem "Poll Messages"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    WebBrowser1.Width = Me.Width - 80
End Sub
