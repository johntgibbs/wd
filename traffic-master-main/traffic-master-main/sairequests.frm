VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form saerequests 
   Caption         =   "SAE Requests"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14835
   LinkTopic       =   "Form2"
   ScaleHeight     =   7530
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   2205
      Left            =   0
      TabIndex        =   9
      Top             =   5280
      Width           =   3495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser3 
      Height          =   2295
      Left            =   3480
      TabIndex        =   8
      Top             =   5280
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   4048
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
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2415
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   4260
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
      Location        =   ""
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2415
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   4260
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
      Location        =   ""
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label rowkey 
      Caption         =   "0"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label barkey 
      Caption         =   "..."
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "SAE Requests:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "saerequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grow As Integer

Private Sub build_dai_expected_receipt()
    Dim i As Integer, s As String, t As String, bc As String
    Dim d As daiexprct
    d.action = "ADD"
    d.sOrderID = ""
    d.dExpectedDate = Format(Now, "MM/dd/yyyy hh:mm:ss")
    d.sItem = ""
    d.sLot = ""
    d.fExpectedQuantity = ""
    d.sStoreDestination = ""
    For i = 0 To List3.ListCount - 1
        s = Trim(LCase(List3.List(i)))
        t = Trim(List3.List(i))
        If Left(s, 20) = "update assignmentid=" Then d.sOrderID = Right(t, Len(t) - 20)
        If Left(s, 8) = "barcode:" Then bc = Right(t, Len(t) - 8)
        'If Left(s, 14) = "item quantity=" Then d.fExpectedQuantity = Right(t, Len(t) - 14)
        If Left(s, 3) = "to:" Then d.sStoreDestination = Right(s, 1)
    Next i
    If bc > " " Then
        d.sItem = Trim(Left(bc, 4))
        'd.sLot = Mid(bc, 5, 8)
        i = Val(rowkey.Caption)                                                         'jv032513
        d.sLot = tmtasks.Grid1.TextMatrix(i, 9)                                         'jv032513
        d.fExpectedQuantity = Val(tmtasks.Grid1.TextMatrix(i, 10)) + Val(tmtasks.Grid1.TextMatrix(i, 12))     'jv032513
    End If
    s = d.sOrderID & ":" & d.sItem & ":" & d.sLot & ":" & d.fExpectedQuantity & ":" & d.sStoreDestination
    'MsgBox s
    If d.sStoreDestination = "2" Or d.sStoreDestination = "3" Or d.sStoreDestination = "5" Then
        tmtasks.Text1.Text = Dai_expected_receipt(d)
        Open "c:\jvwork\daiexpectedreceipt.xml" For Output As #1
        Print #1, tmtasks.Text1.Text
        Close #1
        DoEvents
        tmtasks.WebBrowser1.Navigate2 "c:\jvwork\daiexpectedreceipt.xml"
    End If
End Sub

Private Sub build_sae_response(rtype As String)
    Dim i As Integer, s As String, t As String
    Dim p As saeresponsetype
    p.reqid = ""
    p.warehouse = ""
    p.area = ""
    p.func = ""
    p.barcode = ""
    For i = 0 To List1.ListCount - 1
        s = Trim(LCase(List1.List(i)))
        t = Trim(List1.List(i))
        If Left(s, 11) = "request id=" Then p.reqid = Right(t, Len(t) - 11)
        If Left(s, 18) = "request warehouse=" Then p.warehouse = Right(t, Len(t) - 18)
        If Left(s, 13) = "request area=" Then p.area = Right(t, Len(t) - 13)
        If Left(s, 10) = "func type=" Then p.func = Right(t, Len(t) - 10)
        If Left(s, 5) = "func:" Then p.barcode = Right(t, Len(t) - 5)
    Next i
    p.fromloc = p.warehouse
    p.toloc = "Crane " & tmtasks.Grid2.TextMatrix(1, 0)
    p.uom = "Pallet"
    p.qty = "1"
    p.product = tmtasks.Grid1.TextMatrix(grow, 5)
    p.moveid = tmtasks.Grid1.TextMatrix(grow, 0)
    List2.Clear
    List2.AddItem "reqid:" & p.reqid
    List2.AddItem "warehouse:" & p.warehouse
    List2.AddItem "area:" & p.area
    List2.AddItem "func:" & p.func
    List2.AddItem "barcode:" & p.barcode
    List2.AddItem "fromloc:" & p.fromloc
    List2.AddItem "toloc:" & p.toloc
    List2.AddItem "uom:" & p.uom
    List2.AddItem "qty:" & p.qty
    List2.AddItem "product:" & p.product
    List2.AddItem "moveid:" & p.moveid
    If rtype = "MoveSpecific" Then Call sae_xml_response_move(p)
    If rtype = "BuildPallet" Then Call sae_xml_response_buildpallet(p)
End Sub

Public Sub LoadDocument(docfile As String, lname As Control)
    Dim xdoc As MSXML2.DOMDocument60
    Set xdoc = New MSXML2.DOMDocument60
    xdoc.validateOnParse = False
    'If xdoc.Load("C:\jvwork\testsongs.xml") Then
    If xdoc.Load(docfile) Then
    ' The document loaded successfully.
    ' Now do something intersting.
        'List1.AddItem xdoc.documentElement.nodeName
        lname.AddItem xdoc.documentElement.nodeName
        DisplayNode xdoc.childNodes, 0, lname
    Else
        MsgBox " The document failed to load."
        ' See the previous listing for error information.
    End If
End Sub

Public Sub DisplayNode(ByRef Nodes As IXMLDOMNodeList, ByVal Indent As Integer, lname As Control)
    Dim xNode As IXMLDOMNode
    Dim xattr As IXMLDOMAttribute
    Indent = Indent + 2

    For Each xNode In Nodes
        If xNode.nodeType = NODE_TEXT Then
            'List1.AddItem Space$(Indent) & xNode.parentNode.nodeName &
            lname.AddItem Space$(Indent) & xNode.parentNode.nodeName & _
            ":" & xNode.nodeValue
        Else
            'Text2 = Text2 & xNode.parentNode.nodeName & " type: " & xNode.nodeType & vbCrLf
            'Text2 = Text2 & xNode.nodeName & " type: " & xNode.nodeType & vbCrLf
        End If
      
        If xNode.nodeType = 1 Then
            If xNode.Attributes.length > 0 Then
                For i = 0 To xNode.Attributes.length - 1
                    'Text2.Text = Text2.Text & "attr(" & i & "): " & xNode.Attributes(i).nodeName
                    'List1.AddItem xNode.nodeName & " " & xNode.Attributes(i).nodeName &
                    lname.AddItem xNode.nodeName & " " & xNode.Attributes(i).nodeName & _
                    "=" & xNode.Attributes(i).nodeValue
                Next i
            'Else
            '    Text2 = Text2 & xNode.nodeName & vbCrLf
            End If
        End If
        
      If xNode.hasChildNodes Then
         DisplayNode xNode.childNodes, Indent, lname
      End If
   Next xNode
   
End Sub


Private Sub sae_xml_request(p As saerequesttype)
    Dim xmldoc As MSXML2.DOMDocument60
    Dim ProcInstr As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim timeElement As IXMLDOMElement
    Dim funcElement As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim s As String
    'Creating DOM Document object
    Set xmldoc = New MSXML2.DOMDocument60
    'this adds the processing instruction
    'the first line in an XML document

    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmldoc.appendChild ProcInstr
    'Create the root element
    Set rootElement = xmldoc.createElement("Request")
    Set xmldoc.documentElement = rootElement
    'Creating comment node
    Set comElement = xmldoc.createComment("SAE " & p.func & " request.")
    'add the comment node after the root
    rootElement.appendChild comElement
    
    Set att = xmldoc.createAttribute("ID")
    att.Text = p.id
    rootElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("User")
    att.Text = p.userid
    rootElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Warehouse")
    att.Text = p.warehouse
    rootElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Area")
    att.Text = p.area
    rootElement.Attributes.setNamedItem att
    
    
    Set timeElement = xmldoc.createElement("Time")
    timeElement.nodeTypedValue = Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & "Z"
    rootElement.appendChild timeElement
    
    Set funcElement = xmldoc.createElement("Func")
    funcElement.nodeTypedValue = p.barcode
    rootElement.appendChild funcElement
    
    Set att = xmldoc.createAttribute("Type")
    att.Text = p.func
    funcElement.Attributes.setNamedItem att
    
    '----------------------
    'Saving the xml document to c:testWebStudents.xml
    
    s = "c:\jvwork\sae" & p.func & "request.xml"
    xmldoc.Save s
    WebBrowser1.Navigate2 (s)
    DoEvents
    List1.Clear
    Call LoadDocument(s, List1)
End Sub

Private Sub sae_xml_response_move(p As saeresponsetype)
    Dim xmldoc As MSXML2.DOMDocument60
    Dim ProcInstr As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim taskElement As IXMLDOMElement
    Dim moveElement As IXMLDOMElement
    Dim fromElement As IXMLDOMElement
    Dim frolocElement As IXMLDOMElement
    Dim toElement As IXMLDOMElement
    Dim tolocElement As IXMLDOMElement
    Dim itemElement As IXMLDOMElement
    Dim descElement As IXMLDOMElement
    Dim barcElement As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim s As String
    'Creating DOM Document object
    Set xmldoc = New MSXML2.DOMDocument60
    'this adds the processing instruction
    'the first line in an XML document

    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmldoc.appendChild ProcInstr
    'Create the root element
    Set rootElement = xmldoc.createElement("Assignment")
    Set xmldoc.documentElement = rootElement
    'Creating comment node
    Set comElement = xmldoc.createComment("SAE " & p.func & " response.")
    'add the comment node after the root
    rootElement.appendChild comElement
    
    Set att = xmldoc.createAttribute("ReqID")
    att.Text = p.reqid
    rootElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("ID")
    att.Text = ""
    rootElement.Attributes.setNamedItem att
    
    
    Set taskElement = xmldoc.createElement("Task")
    'taskElement.nodeTypedValue = Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & "Z"
    rootElement.appendChild taskElement
    
    Set moveElement = xmldoc.createElement("Move")
    'moveElement.nodeTypedValue = p.barcode
    taskElement.appendChild moveElement
    
    Set att = xmldoc.createAttribute("ID")
    att.Text = p.moveid
    moveElement.Attributes.setNamedItem att
    
    Set fromElement = xmldoc.createElement("From")
    moveElement.appendChild fromElement
    
    Set frolocElement = xmldoc.createElement("Location")
    frolocElement.nodeTypedValue = p.fromloc
    fromElement.appendChild frolocElement
    
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    frolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Slot"
    frolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Barcode")
    att.Text = ""
    frolocElement.Attributes.setNamedItem att
    
    
    
    Set toElement = xmldoc.createElement("To")
    moveElement.appendChild toElement
    
    Set tolocElement = xmldoc.createElement("Location")
    tolocElement.nodeTypedValue = p.toloc
    toElement.appendChild tolocElement
    
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    tolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Slot"
    tolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Barcode")
    att.Text = ""
    tolocElement.Attributes.setNamedItem att
    
    
    Set itemElement = xmldoc.createElement("Item")
    moveElement.appendChild itemElement
    
    Set att = xmldoc.createAttribute("UOM")
    att.Text = p.uom
    itemElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Quantity")
    att.Text = p.qty
    itemElement.Attributes.setNamedItem att
    
    Set descElement = xmldoc.createElement("Desription")
    descElement.nodeTypedValue = p.product
    itemElement.appendChild descElement
    
    Set barcElement = xmldoc.createElement("Barcode")
    barcElement.nodeTypedValue = p.barcode
    itemElement.appendChild barcElement
    
    '----------------------
    'Saving the xml document to c:testWebStudents.xml
    
    s = "c:\jvwork\sae" & p.func & "response.xml"
    xmldoc.Save s
    WebBrowser2.Navigate2 (s)
    'DoEvents
    'List1.Clear
    'Call LoadDocument(s)
    Call sae_xml_update_move(p)
End Sub

Private Sub sae_xml_response_buildpallet(p As saeresponsetype)
    Dim xmldoc As MSXML2.DOMDocument60
    Dim ProcInstr As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim taskElement As IXMLDOMElement
    Dim moveElement As IXMLDOMElement
    Dim fromElement As IXMLDOMElement
    Dim frolocElement As IXMLDOMElement
    Dim toElement As IXMLDOMElement
    Dim tolocElement As IXMLDOMElement
    Dim itemElement As IXMLDOMElement
    Dim descElement As IXMLDOMElement
    Dim barcElement As IXMLDOMElement
    Dim lotElement As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim s As String
    'Creating DOM Document object
    Set xmldoc = New MSXML2.DOMDocument60
    'this adds the processing instruction
    'the first line in an XML document

    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmldoc.appendChild ProcInstr
    'Create the root element
    Set rootElement = xmldoc.createElement("Assignment")
    Set xmldoc.documentElement = rootElement
    'Creating comment node
    Set comElement = xmldoc.createComment("SAE " & p.func & " response.")
    'add the comment node after the root
    rootElement.appendChild comElement
    
    Set att = xmldoc.createAttribute("ReqID")
    att.Text = p.reqid
    rootElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("ID")
    att.Text = ""
    rootElement.Attributes.setNamedItem att
    
    
    Set taskElement = xmldoc.createElement("Task")
    'taskElement.nodeTypedValue = Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss") & "Z"
    rootElement.appendChild taskElement
    
    Set moveElement = xmldoc.createElement("Move")
    'moveElement.nodeTypedValue = p.barcode
    taskElement.appendChild moveElement
    
    Set att = xmldoc.createAttribute("ID")
    att.Text = p.moveid
    moveElement.Attributes.setNamedItem att
    
    Set fromElement = xmldoc.createElement("From")
    moveElement.appendChild fromElement
    
    Set frolocElement = xmldoc.createElement("Location")
    frolocElement.nodeTypedValue = p.fromloc
    fromElement.appendChild frolocElement
    
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    frolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Slot"
    frolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Barcode")
    att.Text = ""
    frolocElement.Attributes.setNamedItem att
    
    
    
    Set toElement = xmldoc.createElement("To")
    moveElement.appendChild toElement
    
    Set tolocElement = xmldoc.createElement("Location")
    tolocElement.nodeTypedValue = p.toloc
    toElement.appendChild tolocElement
    
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    tolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Slot"
    tolocElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Barcode")
    att.Text = ""
    tolocElement.Attributes.setNamedItem att
    
    
    Set itemElement = xmldoc.createElement("Item")
    moveElement.appendChild itemElement
    
    Set att = xmldoc.createAttribute("UOM")
    att.Text = p.uom
    itemElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Quantity")
    att.Text = p.qty
    itemElement.Attributes.setNamedItem att
    
    Set descElement = xmldoc.createElement("Desription")
    descElement.nodeTypedValue = p.product
    itemElement.appendChild descElement
    
    Set barcElement = xmldoc.createElement("Barcode")
    barcElement.nodeTypedValue = p.barcode
    itemElement.appendChild barcElement
    
    Set lotElement = xmldoc.createElement("Lot")
    'lotElement.nodeTypedValue = "Lot1"
    itemElement.appendChild lotElement
    Set att = xmldoc.createAttribute("Code")
    att.Text = Mid(p.barcode, 5, 8)
    lotElement.Attributes.setNamedItem att
    
    
    '----------------------
    'Saving the xml document to c:testWebStudents.xml
    
    s = "c:\jvwork\sae" & p.func & "response.xml"
    xmldoc.Save s
    WebBrowser2.Navigate2 (s)

End Sub

Private Sub sae_xml_update_move(p As saeresponsetype)
    Dim xmldoc As MSXML2.DOMDocument60
    Dim ProcInstr As IXMLDOMProcessingInstruction
    Dim rootElement As IXMLDOMElement
    Dim moveElement As IXMLDOMElement
    Dim fromElement As IXMLDOMElement
    Dim toElement As IXMLDOMElement
    Dim itemElement As IXMLDOMElement
    Dim descElement As IXMLDOMElement
    Dim barcElement As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim s As String
    'Creating DOM Document object
    Set xmldoc = New MSXML2.DOMDocument60
    'this adds the processing instruction
    'the first line in an XML document

    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    xmldoc.appendChild ProcInstr
    'Create the root element
    Set rootElement = xmldoc.createElement("Update")
    Set xmldoc.documentElement = rootElement
    'Creating comment node
    Set comElement = xmldoc.createComment("SAE " & p.func & " update.")
    'add the comment node after the root
    rootElement.appendChild comElement
    
    Set att = xmldoc.createAttribute("AssignmentID")
    att.Text = p.reqid
    rootElement.Attributes.setNamedItem att
    
    Set moveElement = xmldoc.createElement("Move")
    'moveElement.nodeTypedValue = p.barcode
    rootElement.appendChild moveElement
    
    Set att = xmldoc.createAttribute("ID")
    att.Text = p.moveid
    moveElement.Attributes.setNamedItem att
    
    Set timeElement = xmldoc.createElement("DateTime")
    timeElement.nodeTypedValue = Format(Now, "YYYY-MM-DD") & "T" & Format(Now, "hh:mm:ss")
    moveElement.appendChild timeElement
    
    
    Set fromElement = xmldoc.createElement("From")
    fromElement.nodeTypedValue = p.fromloc
    moveElement.appendChild fromElement
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    fromElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Slot"
    fromElement.Attributes.setNamedItem att
    
    Set toElement = xmldoc.createElement("To")
    toElement.nodeTypedValue = p.toloc
    moveElement.appendChild toElement
    Set att = xmldoc.createAttribute("Primary")
    att.Text = "True"
    toElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Type")
    att.Text = "Pallet"
    toElement.Attributes.setNamedItem att
        
    Set itemElement = xmldoc.createElement("Item")
    moveElement.appendChild itemElement
    
    Set att = xmldoc.createAttribute("UOM")
    att.Text = p.uom
    itemElement.Attributes.setNamedItem att
    Set att = xmldoc.createAttribute("Quantity")
    att.Text = p.qty
    itemElement.Attributes.setNamedItem att
    
    Set descElement = xmldoc.createElement("Desription")
    descElement.nodeTypedValue = p.product
    itemElement.appendChild descElement
    
    Set barcElement = xmldoc.createElement("Barcode")
    barcElement.nodeTypedValue = p.barcode
    itemElement.appendChild barcElement
    
    '----------------------
    'Saving the xml document to c:testWebStudents.xml
    
    s = "c:\jvwork\sae" & p.func & "update.xml"
    xmldoc.Save s
    WebBrowser3.Navigate2 (s)
    DoEvents
    List3.Clear
    Call LoadDocument(s, List3)
    DoEvents
    Call build_dai_expected_receipt
End Sub

Private Sub sae_xml_update_buildpallet(p As saeresponsetype)

End Sub

Sub sae_req_bbpallet()
    Dim p As saerequesttype
    p.id = tmtasks.Grid1.TextMatrix(grow, 16)
    p.userid = "James"
    p.warehouse = tmtasks.Grid1.TextMatrix(grow, 1)
    p.area = "Blue Bell Pallet"
    p.func = "MoveSpecific"
    p.barcode = tmtasks.Grid1.TextMatrix(grow, 6)
    Call sae_xml_request(p)
    DoEvents
    Call build_sae_response("MoveSpecific")
End Sub

Sub sae_buildpallet()
    Dim p As saerequesttype
    p.id = tmtasks.Grid1.TextMatrix(grow, 0)
    p.userid = "James"
    p.warehouse = tmtasks.Grid1.TextMatrix(grow, 1)
    p.area = "Build Pallet"
    p.func = "BuildPallet"
    p.barcode = tmtasks.Grid1.TextMatrix(grow, 6)
    Call sae_xml_request(p)
    DoEvents
    Call build_sae_response("BuildPallet")
End Sub

Private Sub barkey_Change()
    grow = Val(rowkey.Caption)
    If grow > 0 Then
        If Combo1 < " " Then
            Combo1.ListIndex = 0
        Else
            Call Combo1_Click
        End If
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1 = "Move Specific" Then sae_req_bbpallet
    If Combo1 = "Build Pallet" Then sae_buildpallet
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Move Specific"
    Combo1.AddItem "Build Pallet"
End Sub

Private Sub Form_Resize()
    'WebBrowser1.Width = Me.Width - 80
    'WebBrowser2.Width = Me.Width - 80
End Sub

