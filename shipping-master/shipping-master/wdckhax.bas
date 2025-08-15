Attribute VB_Name = "checkhax"
Option Explicit
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nsize As Long) As Long

Public Sub check_hax()
    Dim ret As Long, s As String, uname As String, ffolder As String
    Dim appname As String, cfile As String, foldername As String
    Dim lpbuff As String * 25
    ret = GetUserName(lpbuff, 25)
    On Error Resume Next
    uname = LCase(Left(lpbuff, InStr(lpbuff, Chr(0)) - 1))
    ffolder = "f:\user\" & LCase(uname)
    'MsgBox ffolder
    If uname = "jvierus" Then Exit Sub
    If uname = "ceilers" Then Exit Sub
    If uname = "rlhalfmann" Then Exit Sub
    If uname = "bguyton" Then Exit Sub
    appname = app.EXEName
    foldername = LCase(CurDir)
    If Right(foldername, 6) = "wdapps" Then Exit Sub                        'jv102517
    If foldername = "u:\wdapps" Then Exit Sub
    If foldername = "c:\users\wduser\desktop" Then Exit Sub
    If foldername = "c:\documents and settings\wduser\desktop" Then Exit Sub
    If Right(foldername, 14) = "wduser\desktop" Or True = True Then Exit Sub
    If foldername = ffolder Then Exit Sub
    If Right(foldername, Len(ffolder)) = ffolder Then Exit Sub              'jv032017
    'cfile = "\\bbc-01-prodtrk\wd\data\wdhax.txt"
    'cfile = foldername & "\wdhax.txt"
    cfile = "s:\wd\html\images\wdhax.txt"
    'If Len(Dir(cfile)) > 0 Then
        Open cfile For Append As #1
        Write #1, uname; appname; foldername; Format(Now, "M-d-yyyy h:mm:ss am/pm"); ffolder
        Close #1
    'End If
    s = "Invalid access to " & foldername & "\" & appname & " has been detected." & vbCrLf
    s = s & "Your userid, " & uname & ", has been logged for review by the administrator." & vbCrLf
    s = s & vbCrLf
    s = s & "Have a nice day."
    MsgBox s, vbOKOnly + vbCritical, "Restricted access...."
    End
End Sub

Public Function check_version(vtitle As String, dbname As String) As Boolean
    Dim s As String, vdb As adodb.Connection, vds As adodb.Recordset
    's = Format(Now, "yyMMdd")
    's = "JV " & mid(s, 1, 2) & "." & mid(s, 3, 2) & "." & mid(s, 5, 2)
    'On Error GoTo vberror
    Set vdb = CreateObject("ADODB.Connection")
    vdb.Open dbname
    s = "select * from valuelists where listname = 'latestversion'"
    Set vds = vdb.Execute(s)
    If vds.BOF = False Then
        vds.MoveFirst
        'MsgBox vds!listreturn & " " & Format(Now, "yyMMdd")
        If Format(vds!listreturn, "yyMMdd") < Format(Now, "yyMMdd") Then
            If vtitle <> vds!listdisplay Then
                'MsgBox vtitle & " <> " & vds!listdisplay
                check_version = False
            Else
                check_version = True
            End If
        End If
        
    End If
    vds.Close: vdb.Close
    Exit Function
vberror:
    'eno = Err.Number: edesc = Err.Description: Err.Clear
    'Call vb_elog(eno, edesc, "Function", "check_version(" & vtitle & ", " & dbname & ")", form1.userid)
    'If eno = -2147467259 Then
    '    Resume
    'Else
    '    MsgBox edesc, vbOKOnly, "Function: check_version(" & vtitle & ") - Error Number: " & eno
        End
    'End If
End Function
