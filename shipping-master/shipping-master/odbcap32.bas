Attribute VB_Name = "Module1"
'ODBC Declares
'
Declare Function SQLAllocEnv Lib "odbc32.dll" (env As Long) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal env As Long) As Integer
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal env As Long, hdbc As Long) As Integer
Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal Server As String, ByVal serverlen As Integer, ByVal uid As String, ByVal uidlen As Integer, ByVal pwd As String, ByVal pwdlen As Integer) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Declare Function SQLAllocStmt Lib "odbc32.dll" (ByVal hdbc As Long, hstmt As Long) As Integer
Declare Function SQLFreeStmt Lib "odbc32.dll" (ByVal hstmt As Long, ByVal EndOption As Integer) As Integer
Declare Function SQLExecDirect Lib "odbc32.dll" (ByVal hstmt As Long, ByVal sqlString As String, ByVal sqlstrlen As Long) As Integer
Declare Function SQLNumResultCols Lib "odbc32.dll" (ByVal hstmt As Long, NumCols As Integer) As Integer
Declare Function SQLFetch Lib "odbc32.dll" (ByVal hstmt As Long) As Integer
Declare Function SQLGetData Lib "odbc32.dll" (ByVal hstmt As Long, ByVal Col As Integer, ByVal wConvType As Integer, ByVal lpBuf As String, ByVal dwbuflen As Long, lpcbout As Long) As Integer
Declare Function SQLError Lib "odbc32.dll" (ByVal env As Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal SQLState As String, NativeError As Long, ByVal Buffer As String, ByVal Buflen As Integer, OutLen As Integer) As Integer
Declare Function SQLDescribeCol Lib "odbc32.dll" (ByVal hstmt As Long, ByVal icol As Integer, ByVal szColName As Long, ByVal cbColNameMax As Long, ByVal pcbColName As Integer, ByVal pfSqlType As Integer, ByVal pcbColDef As Long, ByVal pibScale As Long, ByVal pbNullable As Integer) As Integer
'
' ODBC Constants
'
Global Const SQL_SUCCESS = 0
Global Const SQL_SUCCESS_WITH_INFO = 1
Global Const SQL_ERROR = -1
Global Const SQL_NO_DATA_FOUND = 100
Global Const SQL_CLOSE = 0
Global Const SQL_DROP = 1
Global Const SQL_MAX_MESSAGE_LENGTH = 512
Global Const SQL_CHAR = 1
'
' Global constant for declaring fixed length buffer variables:
Global Const gblnBUFFERLEN = 256
'
' Windows API constants for message boxes, mousepointers:
Global Const MB_ICONEXCLAMATION = 48
Global Const DEFAULTCURSOR = 0
'
' Global ODBC Environment, database connection, and statement handles:
Global hEnv As Long
Global hdbc As Long
Global hstmt As Long

Function AllocateODBChEnv(hEnv As Long)
    Dim result As Integer
    AllocateODBChEnv = SQL_SUCCESS
    result = SQLAllocEnv(hEnv)
    If result <> SQL_SUCCESS Then
        MsgBox "Cannot allocate environment handle.", MB_ICONEXCLAMATION, "ODBC Error"
        Screen.MousePointer = 0
        AllocateODBChEnv = result
        Exit Function
    End If
End Function

Function ConnectToDataSource(hEnv, hdbc As Long, hstmt As Long, ByVal DataSource As String, ByVal UserId As String, ByVal Password As String) As Integer
    Dim result As Integer
    ConnectToDataSource = SQL_SUCCESS
    result = SQLAllocConnect(hEnv, hdbc)
    If result <> SQL_SUCCESS Then
        MsgBox "Cannot allocate connection handle.", MB_ICONEXCLAMATION, "ODBC Error"
        Screen.MousePointer = 0
        ConnectToDataSource = result
        Exit Function
    End If
    result = SQLConnect(hdbc, DataSource, Len(DataSource), UserId, Len(UserId), Password, Len(Password))
    If result <> SQL_SUCCESS Then
        MsgBox "Cannot establish datasource connection.", MB_ICONEXCLAMATION, "ODBC Error"
        Screen.MousePointer = 0
        ConnectToDataSource = result
        Exit Function
    End If
    result = SQLAllocStmt(hdbc, hstmt)
    If result <> SQL_SUCCESS Then
        MsgBox "Cannot allocate statement handle.", MB_ICONEXCLAMATION, "ODBC Error"
        Screen.MousePointer = 0
        ConnectToDataSource = result
        Exit Function
    End If
End Function

Function DisconnectFromDataSource(hdbc As Long, hstmt As Long) As Integer
'Function returns false if any API call fails; true otherwise
    Dim result As Integer
    DisconnectFromDataSource = True
    If hstmt <> 0 Then
        result = SQLFreeStmt(hstmt, SQL_DROP)
        If result <> SQL_SUCCESS Then DisconnectFromDataSource = False
    End If
    If hdbc <> 0 Then
        result = SQLDisconnect(hdbc)
        If result <> SQL_SUCCESS Then DisconnectFromDataSource = False
    End If
    If hdbc <> 0 Then
        result = SQLFreeConnect(hdbc)
        If result <> SQL_SUCCESS Then DisconnectFromDataSource = False
    End If
End Function

Sub DisplayODBCError(hdbc As Long, hstmt As Long, WindowCaption As String)
    Dim SQLState As String * 16
    Dim ErrorMsg As String * SQL_MAX_MESSAGE_LENGTH
    Dim ErrMsgSize As Integer
    Dim ErrorCode As Long
    Dim ErrorCodeStr As String
    Dim result As Integer
    SQLState = String$(16, 0)
    ErrorMsg = String$(SQL_MAX_MESSAGE_LENGTH - 1, 0)
    Do
        result = SQLError(0, hdbc, hstmt, SQLState, ErrorCode, ErrorMsg, Len(ErrorMsg), ErrMsgSize)
        Screen.MousePointer = 0
        If result = SQL_SUCCESS Or result = SQL_SUCCESS_WITH_INFO Then
            If ErrMsgSize = 0 Then
                MsgBox "SQL_SUCCESS or SQL_SUCCESS_WITH_ERROR No additional information available.", MB_ICONEXCLAMATION, WindowCaption
            Else
                If ErrorCode = 0 Then
                    ErrorCodeStr = ""
                Else
                    ErrorCodeStr = Trim$(Str(ErrorCode)) & ""
                End If
                MsgBox ErrorCodeStr & Left$(ErrorMsg, ErrMsgSize), MB_ICONEXCLAMATION, WindowCaption
            End If
        End If
    Loop Until result <> SQL_SUCCESS
End Sub

Function Execute_Remote_SQL(ByVal sqlx) As Integer
    Dim result As Integer, temp As Integer
    Execute_Remote_SQL = SQL_SUCCESS
    result = SQLExecDirect(hstmt, sqlx, Len(sqlx))
    If result <> SQL_SUCCESS Then
        Call DisplayODBCError(hdbc, hstmt, "SQL Statement Error during " & sqlx)
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        If temp <> SQL_SUCCESS Then
            Screen.MousePointer = 0
            MsgBox "Cannot free statement handle.", MB_ICONEXCLAMATION, "Unexpected ODBC driver function failure."
        End If
        Execute_Remote_SQL = result
        Exit Function
    End If
    result = SQLFreeStmt(hstmt, SQL_CLOSE)
End Function

Function FreeODBChEnv(hEnv As Long) As Integer
'Returns false if unsuccessful; True otherwise
    Dim result As Integer
    FreeODBChEnv = True
    If hEnv <> 0 Then
        result = SQLFreeEnv(hEnv)
        If result <> SQL_SUCCESS Then FreeODBChEnv = False
    End If
End Function

Function LoadControl(ctlname As Control, ByVal query As String, hstmt As Long, ItemDataFill As Integer, Delimiter As String) As Integer
    Dim result As Integer
    Dim temp As Integer
    Dim RowCnt As Integer
    Dim NumCols As Integer
    Dim ColCnt As Integer
    Dim Buffer As String * gblnBUFFERLEN
    Dim ItemText As String
    Dim ItemDataString As String
    Dim OutLen As Long
    LoadControl = SQL_SUCCESS
    If TypeOf ctlname Is ListBox Then
    ElseIf TypeOf ctlname Is ComboBox Then
    ElseIf TypeOf ctlname Is MSFlexGrid Then
    Else
        LoadControl = -3
        Exit Function
    End If
    'Do the query
    result = SQLExecDirect(hstmt, query, Len(query))
    If result <> SQL_SUCCESS Then
        Call DisplayODBCError(hdbc, hstmt, "SQL Statement Error during LoadControl")
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        If (temp <> SQL_SUCCESS) Then
            MsgBox "Cannot free statement handle.", MB_ICONEXCLAMATION, "Unexpected ODBC Driver function failure"
        End If
        LoadControl = result
        Exit Function
    End If
    'check number of columns returned
    result = SQLNumResultCols(hstmt, NumCols)
    If result <> SQL_SUCCESS Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        'screen.MousePointer = DEFAULTCURSOR
        LoadControl = result
        Exit Function
    End If
    'Set return value to SQL_NO_DATA_FOUND if no rows returned:
    If NumCols = 0 Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        LoadControl = SQL_NO_DATA_FOUND
        Exit Function
    End If
    If TypeOf ctlname Is MSFlexGrid Then
        ctlname.Cols = NumCols + 1
        ctlname.Rows = 2
        ctlname.Clear
    End If
    Buffer = String$(gblnBUFFERLEN, 0)
    'Fill in the control
    RowCnt = 0
    Do
        'Get next row
        result = SQLFetch(hstmt)
        If result <> SQL_SUCCESS Then
            If result = SQL_NO_DATA_FOUND Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                If RowCnt > 0 Then
                    Exit Do
                Else
                    LoadControl = result
                    Exit Function
                End If
            Else
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                LoadControl = result
                Exit Function
            End If
        End If
        RowCnt = RowCnt + 1
        If TypeOf ctlname Is MSFlexGrid Then
            ctlname.Row = RowCnt
        End If
        ItemText = ""
        ItemDataString = ""
        'Get each column
        For ColCnt = 1 To NumCols
            result = SQLGetData(hstmt, ColCnt, SQL_CHAR, Buffer, gblnBUFFERLEN, OutLen) ', "Call to SQLGetData Failer"
            If result <> SQL_SUCCESS Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                LoadControl = result
                Exit Function
            End If
            If TypeOf ctlname Is MSFlexGrid Then
                ctlname.Col = ColCnt
                If OutLen > 0 Then
                    ctlname.Text = Left$(Buffer, OutLen)
                    'set column widths here....
                End If
            Else
                If ItemDataFill And ColCnt = 1 Then
                    If OutLen > 0 Then
                        ItemDataString = Left$(Buffer, OutLen)
                    Else
                        ItemDataString = ""
                    End If
                Else
                    If OutLen > 0 Then
                        If ItemText = "" Then
                            ItemText = Left$(Buffer, OutLen)
                        Else
                            ItemText = ItemText & Delimiter & Left$(Buffer, OutLen)
                        End If
                    Else
                        ItemText = ItemText = ItemText & Delimiter
                    End If
                End If
            End If
        Next ColCnt
        'Add items to control
        If ItemText <> "" Then
            On Error Resume Next
            ctlname.AddItem ItemText
            If Err = 0 Then
                If ItemDataString <> "" Then
                    ctlname.ItemData(ctlname.NewIndex) = Val(ItemDataString)
                End If
            Else
                MsgBox "Result Set too large to fit in control", MB_ICONEXCLAMATION, "LoadControl"
                Exit Do
            End If
            On Error GoTo 0
        End If
        If TypeOf ctlname Is MSFlexGrid Then
            ctlname.Rows = ctlname.Rows + 1
        End If
    Loop
    LoadControl = SQL_SUCCESS
    'screen.MousePointer = 0
End Function

Function LoadGrid(ctlname As Control, ByVal query As String, hstmt As Long, ItemDataFill As Integer, Delimiter As String) As Integer
    Dim result As Integer
    Dim temp As Integer
    Dim RowCnt As Long
    Dim NumCols As Integer
    Dim ColCnt As Integer
    Dim Buffer As String * gblnBUFFERLEN
    Dim OutLen As Long
    Dim i As Integer
    LoadGrid = SQL_SUCCESS
    If TypeOf ctlname Is MSFlexGrid Then
    Else
        LoadGrid = -3
        Exit Function
    End If
    'Do the query
    result = SQLExecDirect(hstmt, query, Len(query))
    If result <> SQL_SUCCESS Then
        Call DisplayODBCError(hdbc, hstmt, "SQL Statement Error during LoadGrid")
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        If (temp <> SQL_SUCCESS) Then
            MsgBox "Cannot free statement handle.", MB_ICONEXCLAMATION, "Unexpected ODBC Driver function failure"
        End If
        LoadGrid = result
        Exit Function
    End If
    'check number of columns returned
    result = SQLNumResultCols(hstmt, NumCols)
    If result <> SQL_SUCCESS Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        LoadGrid = result
        Exit Function
    End If
    'Set return value to SQL_NO_DATA_FOUND if no rows returned:
    If NumCols = 0 Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        LoadGrid = SQL_NO_DATA_FOUND
        Exit Function
    End If
    ctlname.Clear
    ctlname.Cols = NumCols
    ctlname.Rows = 2
    Buffer = String$(gblnBUFFERLEN, 0)
    'Fill in the control
    RowCnt = 0
    Do
        'Get next row
        result = SQLFetch(hstmt)
        If result <> SQL_SUCCESS Then
            If result = SQL_NO_DATA_FOUND Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                If RowCnt > 0 Then
                    Exit Do
                Else
                    LoadGrid = result
                    Exit Function
                End If
            Else
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                LoadGrid = result
                Exit Function
            End If
        End If
        RowCnt = RowCnt + 1
        'Get each column
        For ColCnt = 1 To NumCols
            result = SQLGetData(hstmt, ColCnt, SQL_CHAR, Buffer, gblnBUFFERLEN, OutLen) ', "Call to SQLGetData Failer"
            If result <> SQL_SUCCESS Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                LoadGrid = result
                Exit Function
            End If
            If OutLen > 0 Then ctlname.TextMatrix(RowCnt, ColCnt - 1) = Left$(Buffer, OutLen)
        Next ColCnt
        'Add items to control
        ctlname.Rows = ctlname.Rows + 1
    Loop
    LoadGrid = SQL_SUCCESS
End Function

Sub print_grid(gname As Control, r1 As Integer, r2 As Integer, rtitle As String)
    Dim i As Integer, k As Integer, j As Integer
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long
    
    xs = 0: xe = xs
    For i = 0 To gname.Cols - 1
        If gname.ColWidth(i) > 10 Then xe = xe + gname.ColWidth(i)
    Next i
    If xe > 11600 Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print rtitle
    Printer.Print Format(Now, "mmmm d, yyyy")

    Printer.FontSize = 8
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2 + 1
        ye = j * 240 + 1440
        Printer.Line (xs, ye)-(xe, ye)
        j = j + 1
    Next i
    Printer.DrawWidth = 1
    Printer.FontBold = False
    xm = xs + 100
    For k = 0 To gname.Cols - 1
        If gname.ColWidth(k) > 10 Then
            Printer.PSet (xm, 1230)
            Printer.Print gname.TextMatrix(0, k)
            xm = xm + gname.ColWidth(k)
        End If
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = 0 To gname.Cols - 1
            If gname.ColWidth(k) > 10 Then
                Printer.PSet (xm, j * 240 + 1230)
                Printer.Print gname.TextMatrix(i, k)
                xm = xm + gname.ColWidth(k)
            End If
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = 0 To gname.Cols - 1
        If gname.ColWidth(i) > 10 Then
            Printer.Line (xm, ys)-(xm, ye)
            xm = xm + gname.ColWidth(i)
        End If
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.EndDoc
End Sub

Function add_grid(ctlname As Control, ByVal query As String, rcnt As Long) As Integer
    Dim result As Integer
    Dim temp As Integer
    Dim RowCnt As Long
    Dim NumCols As Integer
    Dim ColCnt As Integer
    Dim Buffer As String * gblnBUFFERLEN
    Dim OutLen As Long
    Dim i As Integer
    add_grid = SQL_SUCCESS
    If TypeOf ctlname Is MSFlexGrid Then
    Else
        add_grid = -3
        Exit Function
    End If
    'Do the query
    result = SQLExecDirect(hstmt, query, Len(query))
    If result <> SQL_SUCCESS Then
        Call DisplayODBCError(hdbc, hstmt, "SQL Statement Error during LoadGrid")
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        MsgBox query
        If (temp <> SQL_SUCCESS) Then
            MsgBox "Cannot free statement handle.", MB_ICONEXCLAMATION, "Unexpected ODBC Driver function failure"
        End If
        add_grid = result
        Exit Function
    End If
    'check number of columns returned
    result = SQLNumResultCols(hstmt, NumCols)
    If result <> SQL_SUCCESS Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        add_grid = result
        Exit Function
    End If
    'Set return value to SQL_NO_DATA_FOUND if no rows returned:
    If NumCols = 0 Then
        temp = SQLFreeStmt(hstmt, SQL_CLOSE)
        add_grid = SQL_NO_DATA_FOUND
        Exit Function
    End If
    'ctlname.Clear
    'ctlname.Cols = NumCols
    'ctlname.Rows = rcnt
    Buffer = String$(gblnBUFFERLEN, 0)
    'Fill in the control
    RowCnt = rcnt
    Do
        'Get next row
        result = SQLFetch(hstmt)
        If result <> SQL_SUCCESS Then
            If result = SQL_NO_DATA_FOUND Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                If RowCnt > 0 Then
                    Exit Do
                Else
                    add_grid = result
                    Exit Function
                End If
            Else
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                add_grid = result
                Exit Function
            End If
        End If
        RowCnt = RowCnt + 1
        'Get each column
        For ColCnt = 1 To NumCols
            result = SQLGetData(hstmt, ColCnt, SQL_CHAR, Buffer, gblnBUFFERLEN, OutLen) ', "Call to SQLGetData Failer"
            If result <> SQL_SUCCESS Then
                temp = SQLFreeStmt(hstmt, SQL_CLOSE)
                add_grid = result
                Exit Function
            End If
            If OutLen > 0 Then ctlname.TextMatrix(RowCnt, ColCnt - 1) = Left$(Buffer, OutLen)
        Next ColCnt
        'Add items to control
        ctlname.Rows = ctlname.Rows + 1
    Loop
    add_grid = SQL_SUCCESS
End Function

