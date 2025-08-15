Attribute VB_Name = "GetUser"
Option Explicit
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nsize As Long) As Long

Public Sub check_hax()
    Dim ret As Long, s As String, uname As String, ffolder As String
    Dim appname As String, cfile As String, foldername As String
    Dim lpbuff As String * 25
    ret = GetUserName(lpbuff, 25)
    On Error Resume Next
    uname = LCase(Left(lpbuff, InStr(lpbuff, Chr(0)) - 1))
    'uname = "abailey"
    ffolder = "f:\user\" & LCase(uname)
    'MsgBox ffolder
    If uname = "jvierus" Or uname = "rlhalfmann" Then Exit Sub
    appname = App.EXEName
    foldername = LCase(CurDir)
    'foldername = LCase(ffolder)
    'MsgBox foldername
    If Right(foldername, 6) = "wdapps" Then Exit Sub                        'jv102517
    If foldername = "u:\wdapps" Then Exit Sub
    If foldername = "c:\users\wduser\desktop" Then Exit Sub
    If foldername = "c:\documents and settings\wduser\desktop" Then Exit Sub
    If Right(foldername, 14) = "wduser\desktop" Then Exit Sub
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

